"""Leitor de arquivo PGN do Chess.com e exportador para planilha Excel.
Autor: Pedro Santos, feito com ajuda do Claude AI
Junho de 2026"""

import io
import re
from datetime import datetime, timedelta
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import urlopen

import pandas as pd
import streamlit as st

TIME_CONTROL_MAP = {
    "600": "Rápida",
    "1800": "Rápida",
    "900": "Rápida",
    "300": "Blitz",
    "180": "Blitz",
    "120": "Blitz",
    "60": "Bala",
    "30": "Bala",
}

TERMINATION_MAP = [
    # Ordem importa: mais específico primeiro
    ("desistência", "Abandono"),
    ("abandonada", "Abandono"),
    ("abandono", "Abandono"),
    ("abandoned", "Abandono"),
    ("resignation", "Abandono"),
    ("resigned", "Abandono"),
    ("xeque-mate", "Xeque-mate"),
    ("checkmate", "Xeque-mate"),
    ("tempo esgotado", "Tempo"),
    ("time", "Tempo"),
    ("acordo", "Acordo"),
    ("agreement", "Acordo"),
    ("repetição", "Repetição"),
    ("repetition", "Repetição"),
    ("material insuficiente", "Material insuficiente"),
    ("insufficient material", "Material insuficiente"),
    ("afogamento", "Afogamento"),
    ("stalemate", "Afogamento"),
    ("50 moves", "Regra dos 50 lances"),
    ("50-move", "Regra dos 50 lances"),
]

def classify_termination(termination):
    """Classifica a forma de término da partida com base na tag Termination."""
    term_lower = termination.lower()
    for key, val in TERMINATION_MAP:
        if key in term_lower:
            return val
    print(f"Warning: Termination '{termination}' not classified.")
    return "?"

def parse_time(end_time_str, utc_offset_hours):
    """Parse EndTime string, apply UTC offset, return HH:MM string."""
    if not end_time_str:
        return ""
    # Format: "18:32:28 GMT+0000" or "1:13:03 GMT+0000"
    time_part = end_time_str.split(" ")[0]
    try:
        parts = time_part.split(":")
        h, m = int(parts[0]), int(parts[1])
        dt = datetime(2000, 1, 1, h, m) + timedelta(hours=utc_offset_hours)
        return dt.strftime("%H:%M")
    except ValueError:
        # Fallback: return up to first 5 chars but handle short strings
        if len(time_part) >= 5 and time_part[2] == ":":
            return time_part[:5]
        elif ":" in time_part:
            return time_part.rsplit(":", 1)[0]
        return time_part

def parse_month_input(value):
    """Converte um valor como 2025-02, 2025/02 ou 202502 em (ano, mês)."""
    if value is None:
        raise ValueError("Informe um mês válido.")

    raw = str(value).strip()
    if not raw:
        raise ValueError("Informe um mês válido.")

    match = re.fullmatch(r"(\d{4})[-/](\d{1,2})", raw)
    if match:
        year, month = int(match.group(1)), int(match.group(2))
    else:
        match = re.fullmatch(r"(\d{6})", raw)
        if match:
            year, month = int(raw[:4]), int(raw[4:])
        else:
            match = re.fullmatch(r"(\d{1,2})[-/](\d{4})", raw)
            if match:
                month, year = int(match.group(1)), int(match.group(2))
            else:
                raise ValueError("Use um formato como 2025-02 ou 2025/02.")

    if not 1 <= month <= 12:
        raise ValueError("O mês deve estar entre 1 e 12.")

    return year, month


def build_month_range(start_month, end_month):
    """Gera a lista de meses entre o início e o fim, inclusive."""
    start_year, start_month = parse_month_input(start_month)
    end_year, end_month = parse_month_input(end_month)

    start_dt = datetime(start_year, start_month, 1)
    end_dt = datetime(end_year, end_month, 1)

    if start_dt > end_dt:
        raise ValueError("O mês final não pode ser anterior ao mês inicial.")

    current = start_dt
    while current <= end_dt:
        yield current.year, current.month
        if current.month == 12:
            current = datetime(current.year + 1, 1, 1)
        else:
            current = datetime(current.year, current.month + 1, 1)


def fetch_pgn_from_chess_com(username, year, month):
    """Busca o arquivo PGN mensal de um jogador no chess.com."""
    url = f"https://api.chess.com/pub/player/{quote(username)}/games/{year}/{month:02d}/pgn"
    try:
        with urlopen(url, timeout=20) as response:
            return response.read().decode("utf-8", errors="replace")
    except HTTPError as exc:
        if exc.code == 404:
            return None
        raise RuntimeError(f"Falha ao buscar o PGN {year}/{month:02d}: HTTP {exc.code}") from exc
    except URLError as exc:
        raise RuntimeError(f"Falha ao buscar o PGN {year}/{month:02d}: {exc}") from exc


def parse_pgn(pgn_text, username, utc_offset):
    """Função principal para processar o texto PGN e retornar um DataFrame com as partidas."""
    games = re.split(r'\n(?=\[Event )', pgn_text.strip())
    records = []

    for game in games:
        if not game.strip():
            continue

        tags = {}
        for match in re.finditer(r'\[(\w+)\s+"([^"]*)"\]', game):
            tags[match.group(1)] = match.group(2)

        white = tags.get("White", "")
        black = tags.get("Black", "")
        result_tag = tags.get("Result", "")
        termination = tags.get("Termination", "")
        time_control = tags.get("TimeControl", "")
        date_str = tags.get("Date", "")
        end_time_str = tags.get("EndTime", "")

        # Determine player color
        if white.lower() == username.lower():
            cor = "Brancas"
            my_elo = tags.get("WhiteElo", "")
        elif black.lower() == username.lower():
            cor = "Pretas"
            my_elo = tags.get("BlackElo", "")
        else:
            cor = "?"
            my_elo = ""

        # Determine result
        if result_tag == "1-0":
            resultado = "Vitória" if cor == "Brancas" else "Derrota"
        elif result_tag == "0-1":
            resultado = "Vitória" if cor == "Pretas" else "Derrota"
        elif result_tag == "1/2-1/2":
            resultado = "Empate"
        else:
            resultado = "?"

        detalhe = classify_termination(termination)

        # Game type
        base_tc = time_control.split("+")[0]
        tipo = TIME_CONTROL_MAP.get(base_tc, f"Outro ({time_control})")

        # Date
        try:
            date_obj = datetime.strptime(date_str, "%Y.%m.%d")
            data = date_obj.strftime("%d/%m/%Y")
        except ValueError:
            data = date_str

        hora = parse_time(end_time_str, utc_offset)

        # Move count
        movetext = re.sub(r'\[.*?\]', '', game).strip()
        movetext = re.sub(r'\{[^}]*\}', '', movetext)
        move_nums = re.findall(r'\b(\d+)\.', movetext)
        num_lances = 0
        if move_nums:
            num_lances = int(move_nums[-1])
            #last_num = int(move_nums[-1])
            #all_moves = re.findall(r'\d+\.\s*(\S+)(?:\s+(\S+))?', movetext)
            #if all_moves:
                #last_black = all_moves[-1][1] if len(all_moves[-1]) > 1 else ""
                #if not last_black or last_black in ("1-0", "0-1", "1/2-1/2", "*"):
                #    num_lances = (last_num - 1) * 2 + 1
                #else:
                #    num_lances = last_num * 2

        records.append({
            "Tipo": tipo,
            "Data": data,
            "Hora": hora,
            "Cor": cor,
            "Resultado": resultado,
            "Detalhe": detalhe,
            "Lances": num_lances,
            "Rating": my_elo,
        })

    return pd.DataFrame(records)

def main():
    """Função principal para a interface Streamlit."""
    st.set_page_config(page_title="PGN → Planilha", page_icon="♟️", layout="centered")

    st.title("♟️ PGN → Planilha")
    st.caption("Exporte suas partidas do Chess.com para Excel")

    col_user, col_tz = st.columns([2, 1])
    with col_user:
        username = st.text_input("Seu nome de usuário no Chess.com", placeholder="ex: PedroSantosG")
    with col_tz:
        utc_offset = st.number_input(
            "Fuso horário (UTC)",
            min_value=-12,
            max_value=14,
            value=-3,
            step=1,
            help="Ex: Brasília = -3",
        )

    source_mode = st.radio(
        "Como quer carregar as partidas?",
        ["Enviar arquivo PGN", "Buscar automaticamente no chess.com"],
        horizontal=True,
    )

    if source_mode == "Enviar arquivo PGN":
        uploaded = st.file_uploader("Envie o arquivo .pgn", type=["pgn", "txt"])

        if uploaded and username.strip():
            pgn_text = uploaded.read().decode("utf-8", errors="replace")

            with st.spinner("Processando partidas..."):
                df = parse_pgn(pgn_text, username.strip(), utc_offset)

            st.success(f"✅ {len(df)} partidas encontradas para **{username}**")
            render_results(df, username)
        elif uploaded and not username.strip():
            st.warning("Preencha o nome de usuário para processar o arquivo.")
    else:
        st.info("Informe o período e o aplicativo buscará os arquivos PGN de cada mês diretamente do chess.com.")
        col_start, col_end = st.columns(2)
        with col_start:
            start_month = st.text_input("Mês inicial", value="2025-02", placeholder="AAAA-MM")
        with col_end:
            end_month = st.text_input("Mês final", value="2025-02", placeholder="AAAA-MM")

        if st.button("Buscar partidas no chess.com", type="primary"):
            if not username.strip():
                st.warning("Preencha o nome de usuário para buscar as partidas.")
            else:
                try:
                    months = list(build_month_range(start_month, end_month))
                except ValueError as exc:
                    st.error(str(exc))
                    months = []

                if months:
                    frames = []
                    with st.spinner("Buscando e processando os arquivos PGN..."):
                        for year, month in months:
                            try:
                                pgn_text = fetch_pgn_from_chess_com(username.strip(), year, month)
                            except RuntimeError as exc:
                                st.warning(str(exc))
                                continue

                            if not pgn_text or not pgn_text.strip():
                                continue

                            month_df = parse_pgn(pgn_text, username.strip(), utc_offset)
                            if month_df.empty:
                                continue

                            month_df = month_df.copy()
                            month_df.insert(0, "Mês", f"{year:04d}-{month:02d}")
                            frames.append(month_df)

                    if frames:
                        df = pd.concat(frames, ignore_index=True)
                        st.success(f"✅ {len(df)} partidas encontradas para **{username}** no período informado")
                        render_results(df, username)
                    else:
                        st.warning("Nenhum arquivo PGN foi encontrado para esse usuário e período.")


def render_results(df, username):
    """Exibe métricas, dados e botão de download a partir de um DataFrame."""
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Vitórias", len(df[df["Resultado"] == "Vitória"]))
    with col2:
        st.metric("Derrotas", len(df[df["Resultado"] == "Derrota"]))
    with col3:
        st.metric("Empates", len(df[df["Resultado"] == "Empate"]))

    st.dataframe(df, width='stretch', hide_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Partidas")
        ws = writer.sheets["Partidas"]
        col_widths = {"A": 12, "B": 12, "C": 8, "D": 10, "E": 10, "F": 22, "G": 8, "H": 8, "I": 10}
        for col, width in col_widths.items():
            if col <= chr(64 + len(df.columns)):
                ws.column_dimensions[col].width = width

    output.seek(0)
    st.download_button(
        label="📥 Baixar planilha Excel",
        data=output,
        file_name=f"partidas_{username}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
    main()
