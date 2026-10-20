import difflib
import io
import re
import unicodedata
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Configurazione Pagina Streamlit
st.set_page_config(
    page_title="Riclassificazione Fatture Trasporti",
    page_icon=":turks_caicos_islands:",
    layout="wide"
)

# ---------------------------------------------------------
# CONFIGURAZIONE AZIENDE
# ---------------------------------------------------------
AZIENDA_ARCONVERT = "Arconvert Spa"
AZIENDA_RITRAMA = "Ritrama Spa"

AZIENDE = {
    AZIENDA_ARCONVERT: {
        # Parola cercata nella colonna "Steeringpartner Company Name" del file 4flow
        "parola_4flow": "ARCONVERT",
        # Con Arconvert il nome fornitore è già nel file transazioni
        "richiede_fornitori": False,
        # Colonne obbligatorie del file transazioni
        "colonne_transazioni": [
            "Posting Date",
            "Document No_",
            "Account_Name",
            "Amount",
            "Description",
            "Document Date",
            "Name",
        ],
        # Fornitori esclusi di default dalla riconciliazione (filtro laterale)
        "esclusi_default": [
            "INTERNATIONAL SERVICES SRL",
            "JAS JET AIR SERVICES SPA",
        ],
    },
    AZIENDA_RITRAMA: {
        "parola_4flow": "RITRAMA",
        "richiede_fornitori": True,
        "colonne_transazioni": [
            "Documento",
            "Data",
            "Data documento",
            "Valore visualizzato conto",
            "Nome conto",
            "Descrizione",
            "Importo nella valuta di dichiarazione",
        ],
        "esclusi_default": [
            "INTERNATIONAL SERVICES SRL",
            "JAS JET AIR SERVICES SPA",
        ],
    },
}

# Colonne obbligatorie dell'anagrafica fornitori (solo Ritrama)
COLONNE_FORNITORI = ["Account fornitore", "Nome"]

# Arconvert: i costi nel file transazioni sono NEGATIVI (es. -88634.14).
# Si inverte il segno per avere i costi positivi come in 4flow e in Ritrama.
SEGNO_IMPORTO_ARCONVERT = -1

# Arconvert: si considerano solo le righe con questo Scenario (se la colonna è presente)
SCENARIO_ARCONVERT = "CONSUNTIVO"

# ---------------------------------------------------------
# SIDEBAR: Selezione Azienda
# ---------------------------------------------------------
st.sidebar.header("🏭 Azienda")
azienda = st.sidebar.radio(
    "Seleziona l'azienda",
    options=list(AZIENDE.keys()),
    index=1,
    horizontal=True,
    label_visibility="collapsed",
)
config_azienda = AZIENDE[azienda]

st.title(f"🚚 Tool Riclassificazione Fatture Trasporti per Competenza – {azienda}")

# ---------------------------------------------------------
# SIDEBAR: Upload File (chiavi diverse per azienda: cambiando azienda si riparte da zero)
# ---------------------------------------------------------
st.sidebar.header("📁 Upload Dati Input")

file_transazioni = st.sidebar.file_uploader(
    f"1. Lista Transazioni {azienda} (Excel/CSV)",
    type=["xlsx", "xls", "csv"],
    key=f"file_transazioni_{azienda}",
)

file_fornitori = None
if config_azienda["richiede_fornitori"]:
    file_fornitori = st.sidebar.file_uploader(
        "2. Lista Fornitori (Excel)",
        type=["xlsx", "xls", "csv"],
        key=f"file_fornitori_{azienda}",
    )
    etichetta_4flow = "3. Estrazione 4flow (Excel)"
else:
    etichetta_4flow = "2. Estrazione 4flow (Excel)"

file_4flow = st.sidebar.file_uploader(
    etichetta_4flow,
    type=["xlsx", "xls"],
    key=f"file_4flow_{azienda}",
)

# ---------------------------------------------------------
# SIDEBAR: Parametri Stanziamento (Mese di BR)
# ---------------------------------------------------------
MESI_IT = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
]

st.sidebar.header("📅 Parametri Stanziamento")

mese_corrente = pd.Timestamp.today().to_period("M")
opzioni_mese_br = list(
    pd.period_range(
        pd.Period(year=mese_corrente.year - 1, month=1, freq="M"),
        pd.Period(year=mese_corrente.year, month=12, freq="M"),
        freq="M",
    )
)

# Default: mese corrente -> il Totale Mese BR arriva fino al mese precedente
mese_br = st.sidebar.selectbox(
    "Mese di BR",
    options=opzioni_mese_br,
    index=opzioni_mese_br.index(mese_corrente),
    format_func=lambda p: f"{MESI_IT[p.month - 1]} {p.year}",
    help=(
        "Il Totale Mese BR somma le differenze da Gennaio fino al mese "
        "PRECEDENTE al Mese di BR (es. Ottobre -> Gennaio-Settembre)."
    ),
)

# Periodo incluso nel Totale Mese BR: da Gennaio al mese precedente al Mese di BR
fine_periodo_br = mese_br - 1
inizio_periodo_br = pd.Period(year=fine_periodo_br.year, month=1, freq="M")

st.sidebar.caption(
    f"Totale Mese BR = somma da **{MESI_IT[inizio_periodo_br.month - 1]} {inizio_periodo_br.year}** "
    f"a **{MESI_IT[fine_periodo_br.month - 1]} {fine_periodo_br.year}**"
)


# ---------------------------------------------------------
# FUNZIONI DI SUPPORTO (lettura file, numeri, date)
# ---------------------------------------------------------
def leggi_file_tabellare(file_caricato):
    """Legge un file Excel o CSV (separatore riconosciuto automaticamente)."""
    nome = file_caricato.name.lower()
    if nome.endswith(".csv"):
        df = pd.read_csv(file_caricato, sep=None, engine="python")
    else:
        df = pd.read_excel(file_caricato)
    df.columns = df.columns.astype(str).str.strip()
    return df


def converti_numero_it(valore):
    """Converte un valore in float gestendo i numeri salvati come testo in formato italiano.

    - Numeri già numerici: restituiti così come sono
    - "1.234,56" -> 1234.56 (punto migliaia, virgola decimale)
    - "3.058"    -> 3058    (solo punto migliaia)
    - "905.35"   -> 905.35  (punto decimale standard)
    """
    if pd.isna(valore):
        return 0.0
    if isinstance(valore, (int, float)):
        return float(valore)

    testo = str(valore).strip().replace(" ", "").replace("€", "")
    if testo == "":
        return 0.0

    if "," in testo:
        testo = testo.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", testo):
        testo = testo.replace(".", "")

    try:
        return float(testo)
    except ValueError:
        return 0.0


def converti_data_it(serie):
    """Converte date in formato italiano gg/mm/aaaa (con o senza ora) o date Excel native."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie

    testo = serie.astype(str).str.strip()
    date = pd.to_datetime(testo, format="%d/%m/%Y %H:%M", errors="coerce")
    date = date.fillna(pd.to_datetime(testo, format="%d/%m/%Y", errors="coerce"))

    mancanti = date.isna() & serie.notna()
    if mancanti.any():
        date[mancanti] = pd.to_datetime(
            serie[mancanti], dayfirst=True, errors="coerce"
        )
    return date


def classifica_tipo_trasporto(nome_conto):
    """Nazionali / Estero / Altro in base al nome del conto."""
    testo = str(nome_conto).upper()
    if "ESTER" in testo:
        return "Estero"
    if "NAZIONAL" in testo or "ITALIA" in testo:
        return "Nazionali"
    return "Altro"


# ---------------------------------------------------------
# FUNZIONI DI ELABORAZIONE LOGICA (CONTABILITÀ)
# ---------------------------------------------------------
def calcola_mese_calcolo(data):
    """Step 3: Determina il Mese Calcolo in base alla data di registrazione (Cutoff giorno 3).

    - Se giorno >= 3: Mese Corrente
    - Se giorno < 3: Mese Precedente
    """
    if pd.isna(data):
        return pd.NaT

    if data.day >= 3:
        return data.to_period("M")
    else:
        return (data - pd.DateOffset(months=1)).to_period("M")


def completa_elaborazione(df_merged):
    """Step comuni a tutte le aziende (2, 3, 4) sul DataFrame con colonne standard:
    Documento, Data, Data documento, Nome, Nome conto, Descrizione,
    Importo nella valuta di dichiarazione.
    """
    # ---------------------------------------------------------
    # STEP 2 & 2.bis: Filtro Esclusione "Trasporti mm.yyyy" e "Trasporti mm.yy"
    # ---------------------------------------------------------
    pattern_esclusione = r"(?i)trasporti\s+\d{1,2}\.\d{2,4}"
    df_filtered = df_merged[
        ~df_merged["Descrizione"]
        .astype(str)
        .str.contains(pattern_esclusione, regex=True, na=False)
    ].copy()

    # ---------------------------------------------------------
    # STEP 3: Mese Calcolo
    # ---------------------------------------------------------
    df_filtered["Data"] = pd.to_datetime(df_filtered["Data"], errors="coerce")
    df_filtered["Mese Calcolo"] = df_filtered["Data"].apply(
        calcola_mese_calcolo
    )

    # ---------------------------------------------------------
    # STEP 4: Fix Data Doc Blank & Competenza da Data Doc
    # ---------------------------------------------------------
    df_filtered["Data documento"] = pd.to_datetime(
        df_filtered["Data documento"], errors="coerce"
    )

    # Se DATA DOC è vuota, utilizziamo la data di registrazione (DATA)
    df_filtered["Data documento"] = df_filtered["Data documento"].fillna(
        df_filtered["Data"]
    )

    # Competenza deriva unicamente dal mese della Data Documento
    df_filtered["Competenza"] = df_filtered["Data documento"].dt.to_period("M")

    # Estrazione numero del mese di competenza (1, 2, 3...) per layout simile all'immagine
    df_filtered["Mese_Competenza_Num"] = df_filtered["Data documento"].dt.month
    df_filtered["Mese_Calcolo_Num"] = df_filtered["Mese Calcolo"].dt.month

    # Formattazione in stringa YYYY-MM
    df_filtered["Mese Calcolo Str"] = df_filtered["Mese Calcolo"].astype(str)
    df_filtered["Competenza Str"] = df_filtered["Competenza"].astype(str)

    # Importo numerico e tipo trasporto (Nazionali / Estero / Altro)
    df_filtered["Importo nella valuta di dichiarazione"] = pd.to_numeric(
        df_filtered["Importo nella valuta di dichiarazione"], errors="coerce"
    ).fillna(0.0)
    df_filtered["Tipo Trasporto"] = df_filtered["Nome conto"].apply(
        classifica_tipo_trasporto
    )

    return df_filtered


def elabora_dati_ritrama(df_trans, df_forn):
    """Ritrama (Dynamics): codice fornitore estratto da 'Valore visualizzato conto'
    e nome recuperato dall'anagrafica fornitori."""
    df_trans = df_trans.copy()
    df_forn = df_forn.copy()

    # ---------------------------------------------------------
    # STEP 1: Estrazione Codice Fornitore e Lookup Nome
    # ---------------------------------------------------------
    df_trans["Codice_Fornitore_Estratto"] = (
        df_trans["Valore visualizzato conto"]
        .astype(str)
        .str.extract(r"--(\d+)$")
    )

    df_trans["Codice_Fornitore_Estratto"] = df_trans[
        "Codice_Fornitore_Estratto"
    ].str.strip()
    df_forn["Account fornitore"] = (
        df_forn["Account fornitore"].astype(str).str.strip()
    )

    # Merge con Anagrafica Fornitori per recuperare 'Nome'
    df_merged = df_trans.merge(
        df_forn[["Account fornitore", "Nome"]],
        left_on="Codice_Fornitore_Estratto",
        right_on="Account fornitore",
        how="left",
    )

    df_merged["Nome"] = df_merged["Nome"].fillna("FORNITORE NON TROVATO")

    return completa_elaborazione(df_merged)


def elabora_dati_arconvert(df_trans):
    """Arconvert: il file contiene già il nome fornitore ('Name').
    Le colonne vengono rinominate nel formato standard usato da tutta l'app."""
    df = df_trans.copy()

    # Solo righe di consuntivo (se la colonna Scenario è presente)
    if "Scenario" in df.columns:
        df = df[df["Scenario"].astype(str).str.strip().str.upper() == SCENARIO_ARCONVERT]

    df_std = pd.DataFrame(
        {
            "Documento": df["Document No_"],
            "Data": converti_data_it(df["Posting Date"]),
            "Data documento": converti_data_it(df["Document Date"]),
            "Codice_Fornitore_Estratto": (
                df["Source No_"].astype(str).str.strip()
                if "Source No_" in df.columns
                else ""
            ),
            "Nome": df["Name"],
            "Conto principale": (
                df["G_L Account No_"] if "G_L Account No_" in df.columns else ""
            ),
            "Nome conto": df["Account_Name"],
            "Descrizione": df["Description"],
            "Importo nella valuta di dichiarazione": (
                df["Amount"].apply(converti_numero_it) * SEGNO_IMPORTO_ARCONVERT
            ),
        },
        index=df.index,
    )

    # Colonne aggiuntive mantenute per il dettaglio
    for col in ["Scenario", "Global Dimension 1 Code"]:
        if col in df.columns:
            df_std[col] = df[col]

    nomi = df_std["Nome"].astype(str).str.strip()
    df_std["Nome"] = nomi.where(
        ~nomi.isin(["", "nan", "None"]), "FORNITORE NON TROVATO"
    )

    return completa_elaborazione(df_std)


def crea_pivot_competenza(df, index, index_name):
    """Pivot importi per mese di competenza (1..12). Tabella vuota se non ci sono dati."""
    if df.empty:
        vuota = pd.DataFrame(dtype=float)
        vuota.index.name = index_name
        return vuota
    pivot = pd.pivot_table(
        df,
        values="Importo nella valuta di dichiarazione",
        index=[index],
        columns=["Mese_Competenza_Num"],
        aggfunc="sum",
        fill_value=0.0,
        margins=True,
        margins_name="Grand Total",
    )
    pivot.index.name = index_name
    return pivot


# ---------------------------------------------------------
# FUNZIONI 4FLOW
# ---------------------------------------------------------
# Voci di costo presenti nell'estrazione 4flow (blocco "Cost Type")
COLONNE_COSTO_4FLOW = [
    "Bunker Surcharges",
    "Customs",
    "Dead freight",
    "Extra stop",
    "Freight charges",
    "Fuel surcharge",
    "Others",
    "Rate claim",
    "Standing trailer costs",
    "Taillift fee",
    "Volume deviation",
    "Waiting time at shipper site / receiver site",
    "Warehouse handling",
]

COLONNA_AZIENDA_4FLOW = "Steeringpartner Company Name"


@st.cache_data(show_spinner=False)
def leggi_4flow(file_bytes):
    """Legge l'estrazione 4flow.

    La riga 1 contiene solo l'etichetta "Cost Type" sopra le colonne di costo:
    l'intestazione vera è sulla riga 2 (header=1).
    """
    df = pd.read_excel(io.BytesIO(file_bytes), header=1, dtype=object)
    df.columns = df.columns.astype(str).str.strip()
    return df


def filtra_4flow_azienda(df_4flow, parola_azienda):
    """Tiene solo le righe 4flow dell'azienda selezionata.

    Restituisce (df_filtrato, righe_altre_aziende) oppure (None, 0) se il file
    non contiene l'azienda o non ha la colonna Steeringpartner Company Name.
    """
    if COLONNA_AZIENDA_4FLOW not in df_4flow.columns:
        return None, 0
    aziende = df_4flow[COLONNA_AZIENDA_4FLOW].astype(str).str.upper()
    maschera = aziende.str.contains(parola_azienda, na=False)
    if not maschera.any():
        return None, 0
    return df_4flow[maschera].copy(), int((~maschera).sum())


def elabora_4flow(df_4flow):
    """Somma tutti i costi di trasporto per riga e prepara Mese Processing e Fornitore."""
    df = df_4flow.copy()

    colonne_costo = [c for c in COLONNE_COSTO_4FLOW if c in df.columns]
    if not colonne_costo:
        raise ValueError(
            "Nessuna colonna di costo trovata nel file 4flow "
            "(es. 'Freight charges', 'Fuel surcharge')."
        )

    for col in colonne_costo:
        df[col] = df[col].apply(converti_numero_it)

    df["Totale Costi 4flow"] = df[colonne_costo].sum(axis=1)

    # Processing Date in formato americano M/D/YYYY (es. 1/14/2026)
    data_testo = pd.to_datetime(
        df["Processing Date"].where(
            df["Processing Date"].apply(lambda x: isinstance(x, str))
        ),
        format="%m/%d/%Y",
        errors="coerce",
    )
    data_nativa = pd.to_datetime(
        df["Processing Date"].where(
            ~df["Processing Date"].apply(lambda x: isinstance(x, str))
        ),
        errors="coerce",
    )
    df["Processing Date"] = data_testo.fillna(data_nativa)

    df["Mese Processing"] = df["Processing Date"].dt.to_period("M").astype(str)
    df.loc[df["Processing Date"].isna(), "Mese Processing"] = "DATA MANCANTE"

    df["Carrier Company Name"] = (
        df["Carrier Company Name"]
        .astype(str)
        .str.strip()
        .replace({"nan": "VETTORE NON INDICATO", "": "VETTORE NON INDICATO"})
    )

    return df, colonne_costo


def crea_pivot_4flow(df_4flow_elab):
    """Pivot: Carrier Company (righe) x Mese Processing (colonne), somma di tutti i costi."""
    pivot = pd.pivot_table(
        df_4flow_elab,
        values="Totale Costi 4flow",
        index=["Carrier Company Name"],
        columns=["Mese Processing"],
        aggfunc="sum",
        fill_value=0.0,
        margins=True,
        margins_name="Grand Total",
    )
    pivot.index.name = "Row Labels"
    pivot.columns.name = None
    return pivot


def crea_pivot_4flow_voci(df_4flow_elab, colonne_costo):
    """Controllo: Voce di costo (righe) x Mese Processing (colonne)."""
    pivot = (
        df_4flow_elab.groupby("Mese Processing")[colonne_costo]
        .sum()
        .T
    )
    pivot = pivot.loc[pivot.abs().sum(axis=1) > 0]
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot.loc["Grand Total"] = pivot.sum(axis=0)
    pivot.index.name = "Voce di costo"
    pivot.columns.name = None
    return pivot


# ---------------------------------------------------------
# FUNZIONI RICONCILIAZIONE 4FLOW vs CONTABILITÀ
# ---------------------------------------------------------
# Nomi "tecnici" che non devono mai essere abbinati
NOMI_DA_ESCLUDERE = {"FORNITORE NON TROVATO", "VETTORE NON INDICATO"}

# Forme societarie e parole "rumore" ignorate nel confronto dei nomi
FORME_SOCIETARIE = {
    "SRL", "SRLS", "SPA", "SAPA", "SNC", "SAS", "SOCIETA", "BENEFIT",
    "UNIPERSONALE", "SOCIO", "UNICO", "GMBH", "KG", "AG", "CO", "LTD",
    "LIMITED", "SA", "SL", "SLU", "BV", "NV", "UAB", "DOO", "SRO", "ZOO",
    "OY", "AB", "INC", "LLC", "SARL", "EOOD", "OOD", "KFT", "ZRT",
}

# Formato mese atteso nelle colonne: YYYY-MM
PATTERN_MESE = r"^\d{4}-\d{2}$"

# Gruppi fornitori: parola chiave (nel nome normalizzato) -> nome del gruppo.
# Applicati sia ai vettori 4flow sia ai fornitori di contabilità, PRIMA dell'abbinamento.
# L'ordine conta: vale la prima parola chiave trovata.
GRUPPI_FORNITORI = {
    "DACHSER": "FERCAM",
    "FERCAM": "FERCAM",
    "EWALS": "EWALS",
    "GIRTEKA": "GIRTEKA"
}


def normalizza_nome_fornitore(nome):
    """Crea una chiave di confronto per i nomi fornitore.

    Esempio: "RABEN ITALY S.R.L." e "Raben Italy Srl" -> "RABEN ITALY"
    """
    if pd.isna(nome):
        return ""

    # Rimozione accenti / caratteri non latini e maiuscolo
    testo = (
        unicodedata.normalize("NFKD", str(nome))
        .encode("ascii", "ignore")
        .decode("ascii")
        .upper()
    )
    testo = testo.replace("&", " AND ")
    testo = testo.replace(".", "")                 # S.P.A. -> SPA
    testo = re.sub(r"[^A-Z0-9]+", " ", testo)      # punteggiatura -> spazio
    testo = re.sub(r"\s+", " ", testo).strip()

    # Unisce lettere singole separate da spazi: "S R L" -> "SRL"
    testo = re.sub(
        r"\b(?:[A-Z] )+[A-Z]\b",
        lambda m: m.group(0).replace(" ", ""),
        testo,
    )

    parole = [p for p in testo.split(" ") if p and p not in FORME_SOCIETARIE]
    chiave = " ".join(parole)
    return chiave if chiave else testo


def raggruppa_fornitore(nome):
    """Restituisce il nome del gruppo se il fornitore contiene una parola chiave
    di GRUPPI_FORNITORI, altrimenti il nome originale.

    Esempio: "DACHSER & FERCAM ITALIA SRL", "Dachser Spain, SA", "FERCAM SPA" -> "FERCAM"
    """
    if pd.isna(nome) or str(nome) in NOMI_DA_ESCLUDERE:
        return nome
    chiave = normalizza_nome_fornitore(nome)
    for parola, gruppo in GRUPPI_FORNITORI.items():
        if parola in chiave:
            return gruppo
    return nome


def tabella_gruppi_applicati(df_4flow_ric, df_cont_ric):
    """Elenco dei nomi originali confluiti in un gruppo, per controllo."""
    righe = []
    for fonte, df, col_orig in [
        ("4flow", df_4flow_ric, "Carrier Company Name"),
        ("Contabilità", df_cont_ric, "Nome"),
    ]:
        coppie = df[[col_orig, "Fornitore Ric"]].drop_duplicates()
        coppie = coppie[coppie[col_orig].astype(str) != coppie["Fornitore Ric"].astype(str)]
        for _, r in coppie.iterrows():
            righe.append(
                {"Gruppo": r["Fornitore Ric"], "Fonte": fonte, "Nome originale": r[col_orig]}
            )
    if not righe:
        return pd.DataFrame(columns=["Gruppo", "Fonte", "Nome originale"])
    return pd.DataFrame(righe).sort_values(["Gruppo", "Fonte", "Nome originale"])


def suggerisci_mapping_fornitori(nomi_4flow, nomi_contabilita, totali_4flow):
    """Propone l'abbinamento Vettore 4flow -> Fornitore Contabilità.

    1. Nome normalizzato identico          -> "Automatico"
    2. Nome contenuto nell'altro / simile  -> "Suggerito - verificare"
    3. Nessun abbinamento                  -> "Non abbinato"
    """
    chiavi_cont = {}
    for nome in sorted(nomi_contabilita):
        chiave = normalizza_nome_fornitore(nome)
        if chiave:
            chiavi_cont.setdefault(chiave, nome)

    righe = []
    for nome_4flow in sorted(nomi_4flow):
        chiave = normalizza_nome_fornitore(nome_4flow)
        abbinato, metodo = None, "Non abbinato"

        if chiave in chiavi_cont:
            abbinato, metodo = chiavi_cont[chiave], "Automatico"
        elif chiave:
            # Uno contenuto nell'altro (es. "FERCAM" in "FERCAM ITALIA")
            contenuti = [
                k for k in chiavi_cont
                if min(len(k), len(chiave)) >= 5 and (k in chiave or chiave in k)
            ]
            if len(contenuti) == 1:
                abbinato = chiavi_cont[contenuti[0]]
                metodo = "Suggerito - verificare"
            else:
                simili = difflib.get_close_matches(
                    chiave, list(chiavi_cont), n=1, cutoff=0.85
                )
                if simili:
                    abbinato = chiavi_cont[simili[0]]
                    metodo = "Suggerito - verificare"

        righe.append(
            {
                "Vettore 4flow": nome_4flow,
                "Fornitore Contabilità": abbinato,
                "Metodo": metodo,
                "Totale 4flow": float(totali_4flow.get(nome_4flow, 0.0)),
            }
        )

    return pd.DataFrame(righe)


def aggiungi_totali(df):
    """Aggiunge colonna e riga 'Grand Total' a una tabella fornitore x mese."""
    df = df.copy()
    df["Grand Total"] = df.sum(axis=1)
    df.loc["Grand Total"] = df.sum(axis=0)
    return df


# Classificazione della riga in base alla fonte in cui il fornitore è presente
FONTE_ENTRAMBI = "Entrambi"
FONTE_SOLO_4FLOW = "Solo 4flow"
FONTE_SOLO_CONT = "Solo contabilità"
ORDINE_FONTE = {FONTE_ENTRAMBI: 0, FONTE_SOLO_4FLOW: 1, FONTE_SOLO_CONT: 2}
COLORI_FONTE = {
    FONTE_ENTRAMBI: "",
    FONTE_SOLO_4FLOW: "background-color: #DDEBF7",
    FONTE_SOLO_CONT: "background-color: #F8CBAD",
}


def somma_per_fornitore_mese(df, col_fornitore, col_mese, col_valore):
    """Tabella fornitore x mese (somma). Restituisce una tabella vuota se non ci sono dati."""
    if df.empty:
        return pd.DataFrame(dtype=float)
    return (
        df.groupby([col_fornitore, col_mese])[col_valore]
        .sum()
        .unstack(fill_value=0.0)
    )


def crea_riconciliazione(df_contabilita, df_4flow_elab, mapping, fornitori_esclusi=()):
    """Riconciliazione COMPLETA per fornitore e mese: 4flow - Contabilità.

    - 4flow:       somma Totale Costi 4flow per Mese Processing
    - Contabilità: somma Importo nella valuta di dichiarazione per mese di Competenza
    - Righe:       TUTTI i fornitori di entrambe le fonti
                   * abbinati          -> riga unica con il nome di contabilità ("Entrambi")
                   * solo in 4flow     -> contabilità = 0 ("Solo 4flow")
                   * solo contabilità  -> 4flow = 0 ("Solo contabilità")
    - Entrambi i DataFrame devono avere la colonna "Fornitore Ric" (nome dopo i gruppi)
    - fornitori_esclusi: righe rimosse dalla riconciliazione (filtro laterale)

    Restituisce: tabella differenze, tabella 4flow, tabella contabilità, dizionario fornitore -> fonte
    """
    fornitori_esclusi = set(fornitori_esclusi)
    mappa = mapping.dropna(subset=["Fornitore Contabilità"]).copy()
    mappa = mappa[mappa["Fornitore Contabilità"].astype(str).str.strip() != ""]
    dizionario_mapping = dict(
        zip(mappa["Vettore 4flow"], mappa["Fornitore Contabilità"])
    )

    # --- Lato 4flow: vettori abbinati -> nome contabilità, altri -> nome 4flow ---
    df_f = df_4flow_elab[
        df_4flow_elab["Mese Processing"].str.match(PATTERN_MESE, na=False)
    ].copy()
    df_f["Fornitore"] = (
        df_f["Fornitore Ric"].map(dizionario_mapping).fillna(df_f["Fornitore Ric"])
    ).astype(str)
    df_f = df_f[~df_f["Fornitore"].isin(fornitori_esclusi)]
    tab_4flow = somma_per_fornitore_mese(
        df_f, "Fornitore", "Mese Processing", "Totale Costi 4flow"
    )

    # --- Lato Contabilità: tutti i fornitori ---
    df_c = df_contabilita[
        df_contabilita["Competenza Str"].astype(str).str.match(PATTERN_MESE, na=False)
    ].copy()
    df_c["Fornitore"] = df_c["Fornitore Ric"].astype(str)
    df_c = df_c[~df_c["Fornitore"].isin(fornitori_esclusi)].copy()
    df_c["Importo"] = pd.to_numeric(
        df_c["Importo nella valuta di dichiarazione"], errors="coerce"
    ).fillna(0.0)
    tab_cont = somma_per_fornitore_mese(df_c, "Fornitore", "Competenza Str", "Importo")

    # --- Fonte di ogni fornitore ---
    fornitori_4flow = set(tab_4flow.index)
    fornitori_cont = set(tab_cont.index)
    fonti = {}
    for f in fornitori_4flow | fornitori_cont:
        if f in fornitori_4flow and f in fornitori_cont:
            fonti[f] = FONTE_ENTRAMBI
        elif f in fornitori_4flow:
            fonti[f] = FONTE_SOLO_4FLOW
        else:
            fonti[f] = FONTE_SOLO_CONT

    # Ordine righe: Entrambi, Solo 4flow, Solo contabilità; poi alfabetico
    fornitori = sorted(fonti, key=lambda f: (ORDINE_FONTE[fonti[f]], f))
    mesi = sorted(set(tab_4flow.columns) | set(tab_cont.columns))

    tab_4flow = tab_4flow.reindex(index=fornitori, columns=mesi, fill_value=0.0)
    tab_cont = tab_cont.reindex(index=fornitori, columns=mesi, fill_value=0.0)
    tab_diff = tab_4flow - tab_cont

    for tab in (tab_4flow, tab_cont, tab_diff):
        tab.index.name = "Fornitore"
        tab.columns.name = None

    return (
        aggiungi_totali(tab_diff),
        aggiungi_totali(tab_4flow),
        aggiungi_totali(tab_cont),
        fonti,
    )


def aggiungi_colonna_fonte(tab, fonti):
    """Inserisce la colonna 'Fonte' come prima colonna (vuota sulla riga Grand Total)."""
    tab = tab.copy()
    tab.insert(0, "Fonte", [fonti.get(f, "") for f in tab.index])
    return tab


def aggiungi_colonne_br(tab_diff, inizio_periodo, fine_periodo):
    """Aggiunge in coda alla tabella di riconciliazione:

    - Totale Mese BR:      somma delle differenze dei mesi da inizio_periodo a fine_periodo
    - Totale Stanziamento: solo i valori POSITIVI di Totale Mese BR (i negativi diventano 0)
                           Nella riga Grand Total: somma degli stanziamenti dei singoli fornitori
    """
    tab = tab_diff.copy()

    mesi_br = [
        c for c in tab.columns
        if re.match(PATTERN_MESE, str(c))
        and inizio_periodo <= pd.Period(str(c), freq="M") <= fine_periodo
    ]

    righe_fornitori = tab.index != "Grand Total"

    tab["Totale Mese BR"] = tab[mesi_br].sum(axis=1) if mesi_br else 0.0
    tab["Totale Stanziamento"] = tab["Totale Mese BR"].clip(lower=0)
    tab.loc["Grand Total", "Totale Stanziamento"] = tab.loc[
        righe_fornitori, "Totale Stanziamento"
    ].sum()

    return tab


# ---------------------------------------------------------
# EXPORT EXCEL
# ---------------------------------------------------------
def applica_stile_excel_pivot(ws, titolo_sezione="ITALIA"):
    """Applica la formattazione visuale identica allo screenshot Excel.

    - Riga 1: Banner Verde con Titolo
    - Riga 3/4: Header azzurro (#9FC5E8)
    - Formattazione numeri: #,##0 (interi)
    - Riga Finale Totali in azzurro
    """
    fill_verde = PatternFill(
        start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"
    )
    fill_azzurro = PatternFill(
        start_color="9FC5E8", end_color="9FC5E8", fill_type="solid"
    )
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # Inserisci 3 righe in alto per far spazio all'intestazione
    ws.insert_rows(1, amount=3)

    # Titolo Sezione A1
    ws["A1"] = titolo_sezione
    ws["A1"].font = font_bold
    ws["A1"].fill = fill_verde
    ws["A1"].alignment = align_center

    # Label misura in A3 e B3
    ws["A3"] = "Sum of Importo nella valuta di dichiarazione"
    ws["A3"].font = font_bold
    ws["B3"] = "Column Labels"
    ws["B3"].font = font_bold

    max_col = ws.max_column
    max_row = ws.max_row

    # Formattazione Header Pivot (Riga 4)
    for col in range(1, max_col + 1):
        cell = ws.cell(row=4, column=col)
        cell.fill = fill_azzurro
        cell.font = font_bold
        if col > 1:
            cell.alignment = align_center
        else:
            cell.alignment = align_left

    # Formattazione Dati e Totali
    for r in range(5, max_row + 1):
        is_grand_total_row = r == max_row
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = font_bold if is_grand_total_row else font_regular

            if is_grand_total_row:
                cell.fill = fill_azzurro

            if c > 1:
                # Formato numero intero con separatore migliaia come in foto
                cell.number_format = "#,##0"
                cell.alignment = align_right
            else:
                cell.alignment = align_left

    # Regolazione automatica larghezza colonne
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)


def genera_excel_export(
    pivot_estero,
    pivot_nazionali,
    pivot_registrazione_comp,
    df_dettaglio,
    pivot_4flow=None,
    tab_riconciliazione=None,
    mapping_fornitori=None,
):
    """Genera un file Excel in memoria formattato ed esteticamente coerente con lo screenshot."""
    output = io.BytesIO()

    # Usiamo direttamente pd.ExcelWriter
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. Scrittura delle 3 Pivot Table nei rispettivi Fogli
        pivot_estero.to_excel(writer, sheet_name="Estero")
        pivot_nazionali.to_excel(writer, sheet_name="Nazionali")
        pivot_registrazione_comp.to_excel(writer, sheet_name="Reg_vs_Competenza")

        # 1.bis Pivot 4flow (solo se il file è stato caricato)
        if pivot_4flow is not None:
            pivot_4flow.to_excel(writer, sheet_name="4flow")

        # 1.ter Riconciliazione 4flow - Contabilità e mapping fornitori
        if tab_riconciliazione is not None:
            tab_riconciliazione.to_excel(writer, sheet_name="Riconciliazione")
        if mapping_fornitori is not None:
            mapping_fornitori.to_excel(
                writer, sheet_name="Mapping Fornitori", index=False
            )

        # 2. Scrittura Dettaglio Transazioni
        df_export = df_dettaglio.copy()
        if "Data" in df_export.columns:
            df_export["Data"] = df_export["Data"].dt.strftime("%Y-%m-%d")
        if "Data documento" in df_export.columns:
            df_export["Data documento"] = df_export[
                "Data documento"
            ].dt.strftime("%Y-%m-%d")

        df_export.to_excel(writer, sheet_name="Dettaglio Transazioni", index=False)

        # 3. Accesso al workbook openpyxl dal writer per applicare lo stile grafico
        wb = writer.book

        applica_stile_excel_pivot(wb["Estero"], titolo_sezione="ESTERO")
        applica_stile_excel_pivot(wb["Nazionali"], titolo_sezione="ITALIA")
        applica_stile_excel_pivot(
            wb["Reg_vs_Competenza"],
            titolo_sezione="ANALISI REGISTRAZIONE VS COMPETENZA",
        )
        if pivot_4flow is not None:
            applica_stile_excel_pivot(
                wb["4flow"],
                titolo_sezione="4FLOW - COSTI PER MESE PROCESSING E VETTORE",
            )
            wb["4flow"]["A3"] = "Sum of Totale Costi 4flow"

        if tab_riconciliazione is not None:
            applica_stile_excel_pivot(
                wb["Riconciliazione"],
                titolo_sezione="RICONCILIAZIONE 4FLOW - CONTABILITÀ",
            )
            wb["Riconciliazione"]["A3"] = "Differenza 4flow - Contabilità"

        header_fill = PatternFill(
            start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"
        )

        if mapping_fornitori is not None:
            ws_mapping = wb["Mapping Fornitori"]
            for cell in ws_mapping[1]:
                cell.font = Font(bold=True)
                cell.fill = header_fill
            for col in ws_mapping.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws_mapping.column_dimensions[
                    get_column_letter(col[0].column)
                ].width = max(max_len + 3, 12)

        # Stile Intestazione Dettaglio Transazioni
        ws_dettaglio = wb["Dettaglio Transazioni"]
        for cell in ws_dettaglio[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill

    return output.getvalue()


# ---------------------------------------------------------
# INTERFACCIA UTENTE & ISTRUZIONI INIZIALI
# ---------------------------------------------------------
file_pronti = file_transazioni is not None and (
    file_fornitori is not None or not config_azienda["richiede_fornitori"]
)

if not file_pronti:
    st.info(
        f"📌 **{azienda}: per il corretto funzionamento dell'applicazione, assicurati di "
        "scaricare e caricare i file richiesti seguendo queste indicazioni:**"
    )

    istruzioni_4flow = f"""
        ### 🚛 File 4flow (opzionale)
        1. Scaricare l'estrazione dei costi di trasporto da **4flow**.
        2. Il file deve contenere almeno le colonne:
           * `Processing Date`
           * `Carrier Company Name`
           * `Steeringpartner Company Name` (deve contenere **{config_azienda['parola_4flow']}**)
           * le voci di costo (`Freight charges`, `Fuel surcharge`, `Taillift fee`, ...)
        3. Lasciare la riga **"Cost Type"** sopra l'intestazione come da estrazione standard.
        """

    if azienda == AZIENDA_RITRAMA:
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("""
            ### 📄 File Transazioni (Dynamics Ritrama)
            1. Andare in **Transazioni Giustificativo**.
            2. Impostare le seguenti colonne nella tabella:
               * `Numero giornale di registrazione`
               * `Documento`
               * `Giustificativo`
               * `Data`
               * `Data documento`
               * `Anno chiuso`
               * `Conto principale`
               * `Valore visualizzato conto`
               * `Nome conto`
               * `Descrizione`
               * `Valuta`
               * `Importo nella valuta della transazione`
               * `Importo`
               * `Importo nella valuta di dichiarazione`
            3. Filtrare la colonna **Conto Principale** con l'opzione **"è uno di"** copiando i codici di conto:
               * `07.02.02.060.02`
               * `07.02.02.060.03`
            4. Filtrare inoltre la data dal **1/1/2026**.
            5. Cliccare con tasto destro e selezionare **"Esporta tutte le righe"**. L'excel verrà poi scaricato.
            """)

        with col2:
            st.markdown("""
            ### 👤 File Anagrafica Fornitori
            1. Selezionare la voce **"Tutti i fornitori"**.
            2. Cliccare con tasto destro sulla colonna e selezionare **"Esporta tutte le righe"**.
            3. L'excel verrà poi scaricato.
            """)

        with col3:
            st.markdown(istruzioni_4flow)

    else:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("""
            ### 📄 File Transazioni (Arconvert)
            Il file deve contenere le colonne:
            * `Scenario` (vengono usate solo le righe **CONSUNTIVO**)
            * `Posting Date` (data registrazione)
            * `Document No_`
            * `G_L Account No_`
            * `Account_Name` (es. *Spese di trasporto su vendite italia*)
            * `Global Dimension 1 Code`
            * `Amount`
            * `Description`
            * `Source No_` (codice fornitore)
            * `Document Date`
            * `Name` (nome fornitore)

            Non è necessario caricare l'anagrafica fornitori.
            """)

        with col2:
            st.markdown(istruzioni_4flow)

else:
    # ---------------------------------------------------------
    # LETTURA E CONTROLLO FILE DI INPUT
    # ---------------------------------------------------------
    errore_input = None
    try:
        df_trans = leggi_file_tabellare(file_transazioni)
        df_forn = (
            leggi_file_tabellare(file_fornitori)
            if config_azienda["richiede_fornitori"]
            else None
        )
    except Exception as e:
        errore_input = f"❌ Errore durante la lettura dei file: {e}"

    if errore_input:
        st.error(errore_input)
        st.stop()

    # Controllo colonne transazioni
    colonne_mancanti = [
        c for c in config_azienda["colonne_transazioni"] if c not in df_trans.columns
    ]
    if colonne_mancanti:
        st.error(f"❌ Errore, carica le transazioni di Dynamics {azienda}")
        st.caption("Colonne mancanti nel file caricato: " + ", ".join(colonne_mancanti))
        st.stop()

    # Controllo colonne anagrafica fornitori (solo Ritrama)
    if config_azienda["richiede_fornitori"]:
        colonne_mancanti_forn = [c for c in COLONNE_FORNITORI if c not in df_forn.columns]
        if colonne_mancanti_forn:
            st.error(f"❌ Errore, carica l'anagrafica fornitori di Dynamics {azienda}")
            st.caption(
                "Colonne mancanti nel file caricato: " + ", ".join(colonne_mancanti_forn)
            )
            st.stop()

    # Controllo file 4flow: deve contenere l'azienda selezionata
    df_4flow_azienda = None
    if file_4flow is not None:
        try:
            df_4flow_letto = leggi_4flow(file_4flow.getvalue())
            df_4flow_azienda, righe_altre_aziende = filtra_4flow_azienda(
                df_4flow_letto, config_azienda["parola_4flow"]
            )
        except Exception:
            df_4flow_azienda, righe_altre_aziende = None, 0

        if df_4flow_azienda is None:
            st.error(
                f"❌ Errore, carica il file 4flow di {azienda} "
                f"(nessuna riga con '{config_azienda['parola_4flow']}' nella colonna "
                f"'{COLONNA_AZIENDA_4FLOW}')"
            )
        elif righe_altre_aziende:
            st.warning(
                f"⚠️ Il file 4flow contiene anche {righe_altre_aziende:,} righe di altre "
                f"aziende: sono state escluse, vengono usate solo quelle di {azienda}."
            )

    # ---------------------------------------------------------
    # ESECUZIONE ELABORAZIONE
    # ---------------------------------------------------------
    try:
        with st.spinner("Elaborazione dati in corso..."):
            if azienda == AZIENDA_RITRAMA:
                df_elaborato = elabora_dati_ritrama(df_trans, df_forn)
            else:
                df_elaborato = elabora_dati_arconvert(df_trans)

        # Elaborazione 4flow (opzionale)
        df_4flow_elab = None
        pivot_4flow = None
        pivot_4flow_voci = None
        if df_4flow_azienda is not None:
            with st.spinner("Elaborazione file 4flow in corso..."):
                df_4flow_elab, colonne_costo_4flow = elabora_4flow(df_4flow_azienda)
                pivot_4flow = crea_pivot_4flow(df_4flow_elab)
                pivot_4flow_voci = crea_pivot_4flow_voci(
                    df_4flow_elab, colonne_costo_4flow
                )

        st.success("✅ Dati elaborati con successo!")

        # ---------------------------------------------------------
        # PIVOT TABLES (Usando i numeri del mese 1..12 per matchare lo screenshot)
        # ---------------------------------------------------------
        df_estero = df_elaborato[df_elaborato["Tipo Trasporto"] == "Estero"]
        df_nazionali = df_elaborato[df_elaborato["Tipo Trasporto"] == "Nazionali"]
        df_altro = df_elaborato[df_elaborato["Tipo Trasporto"] == "Altro"]

        # 1. Estero
        pivot_estero = crea_pivot_competenza(df_estero, "Nome", "Row Labels")

        # 2. Nazionali
        pivot_nazionali = crea_pivot_competenza(df_nazionali, "Nome", "Row Labels")

        # 3. Data Transazione (Mese Calcolo) vs Data Competenza
        pivot_registrazione_comp = crea_pivot_competenza(
            df_elaborato, "Mese Calcolo Str", "Mese Registrazione"
        )

        # Variabili riconciliazione (valorizzate nella tab dedicata)
        tab_riconciliazione = None
        mapping_fornitori = None

        # Visualizzazione Tab Dashboard
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
            [
                "🇮🇹 Trasporti Nazionali",
                "🌍 Trasporti Estero",
                "🔄 Data Registrazione vs Competenza",
                "🚛 4flow",
                "⚖️ Riconciliazione",
                "📋 Dettaglio Dati Processati",
            ]
        )

        with tab1:
            st.markdown("### Trasporti Su Vendite Nazionali (ITALIA)")
            if not df_altro.empty:
                st.warning(
                    f"⚠️ {len(df_altro)} righe con conto non classificabile come Nazionali o "
                    "Estero (conti: " + ", ".join(sorted(df_altro["Nome conto"].astype(str).unique()))
                    + "). Sono escluse da queste due tab ma incluse nella riconciliazione."
                )
            st.dataframe(
                pivot_nazionali.style.format("{:,.0f}"),
                use_container_width=True,
            )

        with tab2:
            st.markdown("### Trasporti Su Vendite Estero")
            st.dataframe(
                pivot_estero.style.format("{:,.0f}"), use_container_width=True
            )

        with tab3:
            st.markdown("### Analisi Data Registrazione vs Competenza")
            st.dataframe(
                pivot_registrazione_comp.style.format("{:,.0f}"),
                use_container_width=True,
            )

        with tab4:
            st.markdown("### 4flow - Costi di trasporto per Mese Processing e Vettore")
            if pivot_4flow is None:
                st.info(
                    f"Carica l'estrazione 4flow di {azienda} nella sidebar per vedere questa analisi."
                )
            else:
                totale_4flow = df_4flow_elab["Totale Costi 4flow"].sum()
                n_vettori = df_4flow_elab["Carrier Company Name"].nunique()
                k1, k2, k3 = st.columns(3)
                k1.metric("Totale costi 4flow", f"€ {totale_4flow:,.0f}")
                k2.metric("Righe 4flow", f"{len(df_4flow_elab):,}")
                k3.metric("Vettori", f"{n_vettori}")

                n_date_mancanti = (
                    df_4flow_elab["Mese Processing"] == "DATA MANCANTE"
                ).sum()
                if n_date_mancanti:
                    st.warning(
                        f"⚠️ {n_date_mancanti} righe senza Processing Date valida "
                        "(raggruppate nella colonna 'DATA MANCANTE')."
                    )

                st.dataframe(
                    pivot_4flow.style.format("{:,.0f}"),
                    use_container_width=True,
                )

                with st.expander("🔍 Controllo: totale per voce di costo e mese"):
                    st.dataframe(
                        pivot_4flow_voci.style.format("{:,.0f}"),
                        use_container_width=True,
                    )

        with tab5:
            st.markdown("### Riconciliazione 4flow − Contabilità per Fornitore e Mese")
            if df_4flow_elab is None:
                st.info(
                    f"Carica l'estrazione 4flow di {azienda} nella sidebar per calcolare la riconciliazione."
                )
            else:
                st.caption(
                    "Differenza = **4flow** (per Mese Processing) − **Contabilità** "
                    "(per mese di Competenza, da Data documento). "
                    "Valore positivo: trasportato ma non ancora fatturato → da stanziare. "
                    "Valore negativo: fatturato più di quanto risulta in 4flow."
                )

                # Contenitore per la tabella, mostrato sopra all'abbinamento
                area_tabella = st.container()

                # --- Applicazione Gruppi Fornitori (su copie, senza toccare i dati originali) ---
                df_4flow_ric = df_4flow_elab.copy()
                df_4flow_ric["Fornitore Ric"] = df_4flow_ric["Carrier Company Name"].apply(
                    raggruppa_fornitore
                )
                df_cont_ric = df_elaborato.copy()
                df_cont_ric["Fornitore Ric"] = (
                    df_cont_ric["Nome"].astype(str).apply(raggruppa_fornitore)
                )

                # --- Abbinamento Vettori 4flow -> Fornitori Contabilità ---
                nomi_4flow = set(df_4flow_ric["Fornitore Ric"]) - NOMI_DA_ESCLUDERE
                nomi_cont = set(df_cont_ric["Fornitore Ric"].astype(str)) - NOMI_DA_ESCLUDERE
                totali_4flow = df_4flow_ric.groupby("Fornitore Ric")[
                    "Totale Costi 4flow"
                ].sum()

                gruppi_applicati = tabella_gruppi_applicati(df_4flow_ric, df_cont_ric)
                if len(gruppi_applicati):
                    with st.expander(
                        f"🧩 Gruppi fornitori applicati ({gruppi_applicati['Gruppo'].nunique()} gruppi)"
                    ):
                        st.caption(
                            "Nomi originali confluiti in un unico gruppo "
                            "(regole definite in GRUPPI_FORNITORI nel codice)."
                        )
                        st.dataframe(
                            gruppi_applicati,
                            hide_index=True,
                            use_container_width=True,
                            height=500
                        )

                mapping_default = suggerisci_mapping_fornitori(
                    nomi_4flow, nomi_cont, totali_4flow
                )

                with st.expander(
                    "🔗 Abbinamento Vettori 4flow ↔ Fornitori Contabilità (verifica e correggi)",
                    expanded=False,
                ):
                    st.caption(
                        "Gli abbinamenti 'Suggerito - verificare' vanno controllati. "
                        "Per abbinare un vettore non trovato, selezionare il fornitore "
                        "dalla colonna 'Fornitore Contabilità'. Se lasciato vuoto, il vettore "
                        "compare comunque in riconciliazione come 'Solo 4flow'."
                    )
                    mapping_fornitori = st.data_editor(
                        mapping_default,
                        key=f"editor_mapping_{azienda}_{len(nomi_4flow)}_{len(nomi_cont)}",
                        hide_index=True,
                        use_container_width=True,
                        disabled=["Vettore 4flow", "Metodo", "Totale 4flow"],
                        column_config={
                            "Fornitore Contabilità": st.column_config.SelectboxColumn(
                                "Fornitore Contabilità",
                                options=sorted(nomi_cont),
                                required=False,
                            ),
                            "Totale 4flow": st.column_config.NumberColumn(
                                "Totale 4flow", format="%.0f"
                            ),
                        },
                    )

                # --- Filtro laterale: fornitori da escludere dalla riconciliazione ---
                vettori_abbinati = set(
                    mapping_fornitori.dropna(subset=["Fornitore Contabilità"])["Vettore 4flow"]
                )
                opzioni_esclusione = sorted(
                    set(df_cont_ric["Fornitore Ric"].astype(str))
                    | (set(df_4flow_ric["Fornitore Ric"].astype(str)) - vettori_abbinati)
                )
                chiavi_default = {
                    normalizza_nome_fornitore(n) for n in config_azienda["esclusi_default"]
                }
                esclusi_default = [
                    o for o in opzioni_esclusione
                    if normalizza_nome_fornitore(o) in chiavi_default
                ]

                st.sidebar.header("🚫 Fornitori Esclusi")
                fornitori_esclusi = st.sidebar.multiselect(
                    "Cerca e seleziona i fornitori da rimuovere dalla riconciliazione",
                    options=opzioni_esclusione,
                    default=esclusi_default,
                    key=f"fornitori_esclusi_{azienda}",
                )

                (
                    tab_riconciliazione,
                    tab_ric_4flow,
                    tab_ric_cont,
                    fonti_fornitori,
                ) = crea_riconciliazione(
                    df_cont_ric, df_4flow_ric, mapping_fornitori, fornitori_esclusi
                )
                tab_riconciliazione = aggiungi_colonne_br(
                    tab_riconciliazione, inizio_periodo_br, fine_periodo_br
                )
                tab_riconciliazione = aggiungi_colonna_fonte(
                    tab_riconciliazione, fonti_fornitori
                )

                with area_tabella:
                    if tab_riconciliazione.shape[0] <= 1:
                        st.warning(
                            "Nessun dato disponibile per la riconciliazione."
                        )
                    else:
                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric(
                            "Totale 4flow",
                            f"€ {tab_ric_4flow.loc['Grand Total', 'Grand Total']:,.0f}",
                        )
                        k2.metric(
                            "Totale Contabilità",
                            f"€ {tab_ric_cont.loc['Grand Total', 'Grand Total']:,.0f}",
                        )
                        k3.metric(
                            "Differenza 4flow − Contabilità",
                            f"€ {tab_riconciliazione.loc['Grand Total', 'Grand Total']:,.0f}",
                        )
                        k4.metric(
                            f"Totale Stanziamento (BR {MESI_IT[mese_br.month - 1]} {mese_br.year})",
                            f"€ {tab_riconciliazione.loc['Grand Total', 'Totale Stanziamento']:,.0f}",
                        )

                        st.caption(
                            f"**Totale Mese BR**: somma delle differenze da "
                            f"{MESI_IT[inizio_periodo_br.month - 1]} a "
                            f"{MESI_IT[fine_periodo_br.month - 1]} {fine_periodo_br.year}. "
                            "**Totale Stanziamento**: solo i Totale Mese BR positivi "
                            "(i negativi valgono 0)."
                        )
                        if fornitori_esclusi:
                            st.caption(
                                "🚫 **Esclusi dalla riconciliazione** (filtro laterale): "
                                + ", ".join(fornitori_esclusi)
                            )

                        # --- Avviso fornitori presenti in una sola fonte ---
                        solo_cont = sorted(
                            f for f, v in fonti_fornitori.items() if v == FONTE_SOLO_CONT
                        )
                        solo_4flow = sorted(
                            f for f, v in fonti_fornitori.items() if v == FONTE_SOLO_4FLOW
                        )
                        if solo_cont:
                            st.warning(
                                f"🟥 **{len(solo_cont)} fornitori solo in contabilità** "
                                "(fatture presenti, consegne non transitate in 4flow → "
                                "eventuale segnalazione a IT): " + ", ".join(solo_cont)
                            )
                        if solo_4flow:
                            st.info(
                                f"🟦 **{len(solo_4flow)} vettori solo in 4flow** "
                                "(trasporti senza fatture registrate, oppure vettore non "
                                "abbinato: verificare l'abbinamento): " + ", ".join(solo_4flow)
                            )

                        # --- Filtro visualizzazione per Fonte ---
                        fonti_selezionate = st.multiselect(
                            "Mostra fornitori",
                            options=[FONTE_ENTRAMBI, FONTE_SOLO_4FLOW, FONTE_SOLO_CONT],
                            default=[FONTE_ENTRAMBI, FONTE_SOLO_4FLOW, FONTE_SOLO_CONT],
                        )

                        colonne_numeriche = [
                            c for c in tab_riconciliazione.columns if c != "Fonte"
                        ]
                        vista = tab_riconciliazione.drop(index="Grand Total")
                        vista = vista[vista["Fonte"].isin(fonti_selezionate)].copy()
                        etichetta_totale = (
                            "Grand Total" if len(fonti_selezionate) == 3 else "Totale selezione"
                        )
                        riga_totale = vista[colonne_numeriche].sum()
                        vista.loc[etichetta_totale, colonne_numeriche] = riga_totale
                        vista.loc[etichetta_totale, "Fonte"] = ""

                        st.dataframe(
                            vista.style.format("{:,.0f}", subset=colonne_numeriche)
                            .set_properties(
                                subset=["Totale Mese BR", "Totale Stanziamento"],
                                **{"background-color": "#FFF2CC", "font-weight": "bold"},
                            )
                            .apply(
                                lambda s: [COLORI_FONTE.get(v, "") for v in s],
                                subset=["Fonte"],
                            ),
                            use_container_width=True,
                        )

                        with st.expander("🔍 Dettaglio: importi 4flow e Contabilità allineati"):
                            st.markdown("**4flow (per Mese Processing)**")
                            st.dataframe(
                                tab_ric_4flow.style.format("{:,.0f}"),
                                use_container_width=True,
                            )
                            st.markdown("**Contabilità (per mese di Competenza)**")
                            st.dataframe(
                                tab_ric_cont.style.format("{:,.0f}"),
                                use_container_width=True,
                            )

        with tab6:
            st.markdown("### Transazioni Trasformate (Anteprima)")
            cols_preview = [
                c for c in [
                    "Documento",
                    "Data",
                    "Data documento",
                    "Codice_Fornitore_Estratto",
                    "Nome",
                    "Nome conto",
                    "Tipo Trasporto",
                    "Descrizione",
                    "Importo nella valuta di dichiarazione",
                    "Mese Calcolo Str",
                    "Competenza Str",
                ]
                if c in df_elaborato.columns
            ]
            st.dataframe(
                df_elaborato[cols_preview], use_container_width=True
            )

        # ---------------------------------------------------------
        # EXPORT EXCEL (in sidebar, dopo la riconciliazione per includere il mapping)
        # ---------------------------------------------------------
        st.sidebar.subheader("📥 Export Report Excel")
        excel_data = genera_excel_export(
            pivot_estero,
            pivot_nazionali,
            pivot_registrazione_comp,
            df_elaborato,
            pivot_4flow=pivot_4flow,
            tab_riconciliazione=tab_riconciliazione,
            mapping_fornitori=mapping_fornitori,
        )

        st.sidebar.download_button(
            label="💾 Scarica Report Excel Formattato",
            data=excel_data,
            file_name=f"Riclassificazione_Trasporti_{azienda.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione dei file: {e}")