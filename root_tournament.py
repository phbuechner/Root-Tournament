import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import copy
import random
from io import BytesIO
from collections import defaultdict, Counter

# --- Konstanten (Deutsche Bezeichnungen) ---
FACTIONS = [
    "Marquise de Katz", "Baumkronen-Dynastie", "Waldland-Allianz", "Vagabund",
    "Flussvolk-Kompanie", "Echsen-Kult", "Untergrund-Herzogtum", "Krähen-Komplott",
]
MAPS = ["Herbst", "Winter", "Berg", "See"]
TOURNAMENT_POINTS_MAP = {1: 5, 2: 4, 3: 3, 4: 2, 5: 1}

# --- Hilfsfunktionen ---
def initialize_state():
    """Initialisiert den Session State, falls noch nicht geschehen."""
    if 'players' not in st.session_state: st.session_state.players = []
    if 'games' not in st.session_state: st.session_state.games = []
    if 'next_turn_order_names' not in st.session_state: st.session_state.next_turn_order_names = []
    if 'initial_turn_order' not in st.session_state: st.session_state.initial_turn_order = []
    if 'initialized' not in st.session_state: st.session_state.initialized = False
    if 'num_players_input' not in st.session_state: st.session_state.num_players_input = 2
    if 'tournament_finished' not in st.session_state: st.session_state.tournament_finished = False
    if 'simulated_map_index' not in st.session_state: st.session_state.simulated_map_index = 0
    if 'simulated_factions' not in st.session_state: st.session_state.simulated_factions = {}
    if 'simulated_vps' not in st.session_state: st.session_state.simulated_vps = {}
    if 'simulation_triggered' not in st.session_state: st.session_state.simulation_triggered = False

def get_player_data_by_name(name):
    """Gibt die Daten eines Spielers anhand des Namens zurück."""
    for player in st.session_state.players:
        if player['name'] == name: return player
    return None

def calculate_next_turn_order(players):
    """Berechnet die Zugreihenfolge für das nächste Spiel (ab Spiel 2)."""
    if not players: return []
    for p in players:
        p.setdefault('last_vp', 0)
        p.setdefault('total_tp', 0)
    sorted_players = sorted(players, key=lambda p: (p['total_tp'], -p['last_vp']))
    return [p['name'] for p in sorted_players]

def generate_standings_df(players):
    """Erstellt ein Pandas DataFrame für die Rangliste."""
    if not players:
        return pd.DataFrame(columns=['Rang', 'Name', 'Ges. Turnierpkt.', 'Ges. Siegpunkte', 'Siege', 'Letzte Spiel VP', 'Ø Platzierung', 'Gespielte Fraktionen'])

    for p in players:
        p.setdefault('total_vp', 0)
        p.setdefault('wins', 0)
        p.setdefault('total_tp', 0)
        p.setdefault('last_vp', 0)
        p.setdefault('total_placement_sum', 0)
        p.setdefault('games_played', 0)
        p.setdefault('played_factions_str', '')
        p.setdefault('name', 'Unbekannt')
        p.setdefault('id', -1)

    display_players = copy.deepcopy(players)
    num_games = len(st.session_state.get('games', []))
    if num_games > 0:
        for p in display_players:
             p['avg_placement'] = f"{p['total_placement_sum'] / p['games_played']:.2f}" if p['games_played'] > 0 else '-'
    else:
        for p in display_players: p['avg_placement'] = '-'

    df = pd.DataFrame(display_players)
    expected_cols = ['name', 'total_tp', 'total_vp', 'wins', 'last_vp', 'avg_placement', 'played_factions_str']
    if not all(col in df.columns for col in expected_cols):
        st.error("Fehler bei der DataFrame-Erstellung. Nicht alle erwarteten Spalten sind vorhanden.")
        return pd.DataFrame(columns=['Rang', 'Name', 'Ges. Turnierpkt.', 'Ges. Siegpunkte', 'Siege', 'Letzte Spiel VP', 'Ø Platzierung', 'Gespielte Fraktionen'])

    df = df.sort_values(by=['total_tp', 'total_vp'], ascending=False).reset_index(drop=True)
    df['Rang'] = df.index + 1
    df = df[['Rang', 'name', 'total_tp', 'total_vp', 'wins', 'last_vp', 'avg_placement', 'played_factions_str']]
    df.columns = ['Rang', 'Name', 'Ges. Turnierpkt.', 'Ges. Siegpunkte', 'Siege', 'Letzte Spiel VP', 'Ø Platzierung', 'Gespielte Fraktionen']
    return df

def generate_plot_data(games, players):
    """Bereitet Daten für das Plotly-Diagramm vor."""
    plot_data = []
    player_names = [p['name'] for p in players if 'name' in p]
    if not player_names: return pd.DataFrame()

    player_points_over_time = {name: [0] for name in player_names}
    for game in games:
        temp_player_points = {name: player_points_over_time[name][-1] for name in player_names}
        for player_result in game.get('results', []):
            player_name = player_result.get('name')
            tp = player_result.get('tp', 0)
            if player_name in temp_player_points: temp_player_points[player_name] += tp
        for player_name, total_points in temp_player_points.items():
            if player_name in player_points_over_time: player_points_over_time[player_name].append(total_points)

    for player_name, points_list in player_points_over_time.items():
        for game_idx, points in enumerate(points_list):
            plot_data.append({'Spiel': game_idx, 'Spieler': player_name, 'Kumulierte Turnierpunkte': points})
    
    return pd.DataFrame(plot_data) if plot_data else pd.DataFrame()

def calculate_faction_stats(games, factions):
    """Berechnet Statistiken für jede Fraktion."""
    if not games: return pd.DataFrame(columns=['Fraktion', 'Gespielt', 'Siege', 'Ø Siegpunkte', 'Ø Turnierpunkte'])

    faction_data = defaultdict(lambda: {'count': 0, 'total_vp': 0, 'total_tp': 0, 'wins': 0})
    for game in games:
        for result in game.get('results', []):
            faction, vp, tp, rank = result.get('faction'), result.get('vp', 0), result.get('tp', 0), result.get('rank')
            if faction:
                faction_data[faction]['count'] += 1
                faction_data[faction]['total_vp'] += vp
                faction_data[faction]['total_tp'] += tp
                if rank == 1: faction_data[faction]['wins'] += 1

    stats_list = []
    for faction in factions:
        data = faction_data[faction]
        count = data['count']
        if count > 0:
            avg_vp, avg_tp = data['total_vp'] / count, data['total_tp'] / count
            stats_list.append({'Fraktion': faction, 'Gespielt': count, 'Siege': data['wins'], 'Ø Siegpunkte': f"{avg_vp:.2f}", 'Ø Turnierpunkte': f"{avg_tp:.2f}"})
        else:
            stats_list.append({'Fraktion': faction, 'Gespielt': 0, 'Siege': 0, 'Ø Siegpunkte': '-', 'Ø Turnierpunkte': '-'})
    
    df = pd.DataFrame(stats_list)
    return df.sort_values(by='Gespielt', ascending=False).reset_index(drop=True)

def calculate_map_stats(games, maps):
    """Berechnet Statistiken für jede Karte."""
    if not games: return pd.DataFrame(columns=['Karte', 'Gespielt (Spiele)', 'Ø Siegpunkte (Gesamt)'])
    
    map_data = defaultdict(lambda: {'count': 0, 'total_vp': 0, 'player_games': 0})
    for game in games:
        game_map = game.get('map')
        if not game_map: continue
        map_data[game_map]['count'] += 1
        for result in game.get('results', []):
            map_data[game_map]['total_vp'] += result.get('vp', 0)
            map_data[game_map]['player_games'] += 1

    stats_list = []
    for map_name in maps:
        data = map_data[map_name]
        player_games_on_map = data['player_games']
        if player_games_on_map > 0:
            avg_vp = data['total_vp'] / player_games_on_map
            stats_list.append({'Karte': map_name, 'Gespielt (Spiele)': data['count'], 'Ø Siegpunkte (Gesamt)': f"{avg_vp:.2f}"})
        else:
            stats_list.append({'Karte': map_name, 'Gespielt (Spiele)': 0, 'Ø Siegpunkte (Gesamt)': '-'})

    df = pd.DataFrame(stats_list)
    return df.sort_values(by='Gespielt (Spiele)', ascending=False).reset_index(drop=True)

def df_to_excel(df_dict):
    """Exportiert mehrere DataFrames in eine Excel-Datei mit mehreren Blättern."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def recalculate_player_stats_from_games(players, games):
    """Setzt Spielerstatistiken zurück und berechnet sie komplett neu basierend auf der 'games'-Liste."""
    for p in players:
        p.update({'total_tp': 0, 'total_vp': 0, 'wins': 0, 'last_vp': 0, 'played_factions': [], 'played_factions_str': '', 'total_placement_sum': 0, 'games_played': 0})

    for game in sorted(games, key=lambda g: g['game_number']):
        current_num_players = len(game['results'])
        current_points_map = {rank: TOURNAMENT_POINTS_MAP.get(rank, 0) for rank in range(1, current_num_players + 1)}
        for result in game['results']:
            player_data = get_player_data_by_name(result['name'])
            if player_data:
                tp = current_points_map.get(result['rank'], 0)
                player_data['total_tp'] += tp
                player_data['total_vp'] += result['vp']
                player_data['last_vp'] = result['vp']
                if result['faction'] not in player_data['played_factions']:
                    player_data['played_factions'].append(result['faction'])
                player_data['played_factions_str'] = ", ".join(sorted(player_data['played_factions']))
                player_data['total_placement_sum'] += result['rank']
                player_data['games_played'] += 1
                if result['rank'] == 1: player_data['wins'] += 1

def import_from_excel(uploaded_file):
    """Versucht, den Turnierstatus aus einer hochgeladenen Excel-Datei wiederherzustellen."""
    try:
        df_logs = pd.read_excel(uploaded_file, sheet_name='Spielprotokolle')
        required_cols = ["Spiel Nr", "Spieler", "Fraktion", "Siegpunkte (VP)", "Platz", "Turnierpunkte (TP)", "Karte", "Zugreihenfolge (Spiel)"]
        if not all(col in df_logs.columns for col in required_cols):
            st.error("Importfehler: Die Excel-Datei hat nicht das erwartete Format. Wichtige Spalten im 'Spielprotokolle'-Blatt fehlen.")
            return False

        player_names = df_logs['Spieler'].unique().tolist()
        st.session_state.num_players = len(player_names)
        st.session_state.players = [{'id': i, 'name': name} for i, name in enumerate(player_names)]

        reconstructed_games = []
        for game_num, game_df in df_logs.groupby('Spiel Nr'):
            game_log_entry = {
                'game_number': int(game_num), 'map': game_df['Karte'].iloc[0],
                'turn_order': game_df['Zugreihenfolge (Spiel)'].iloc[0].split(' → '),
                'results': []
            }
            for _, row in game_df.iterrows():
                game_log_entry['results'].append({'name': row['Spieler'], 'faction': row['Fraktion'], 'vp': int(row['Siegpunkte (VP)']), 'rank': int(row['Platz']), 'tp': int(row['Turnierpunkte (TP)'])})
            reconstructed_games.append(game_log_entry)
        st.session_state.games = reconstructed_games

        if st.session_state.games:
            game1 = next((g for g in st.session_state.games if g['game_number'] == 1), None)
            if game1: st.session_state.initial_turn_order = game1['turn_order']

        recalculate_player_stats_from_games(st.session_state.players, st.session_state.games)
        st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)
        st.session_state.initialized = True
        st.session_state.tournament_finished = False
        return True
    except Exception as e:
        st.error(f"Ein Fehler ist beim Import aufgetreten: {e}")
        st.error("Stelle sicher, dass du eine gültige, von diesem Tool exportierte Excel-Datei hochlädst.")
        return False

# --- Streamlit App ---
st.set_page_config(page_title="Root Turnier Manager", layout="wide", initial_sidebar_state="expanded")
st.title("🏆 Root Brettspiel Turnier Manager")

initialize_state()

# --- 1. Spieler-Setup (Nur am Anfang) ---
if not st.session_state.initialized:
    col_setup, col_import = st.columns(2)
    with col_setup:
        st.subheader("Neues Turnier starten")
        with st.form("player_form"):
            num_players = st.number_input("Anzahl der Spieler", min_value=2, max_value=10, value=st.session_state.num_players_input, step=1, key="num_players_selector")
            st.session_state.num_players_input = num_players
            st.divider()
            st.write("**Spielernamen eingeben:**")
            player_name_inputs = [st.text_input(f"Name Spieler {i+1}", key=f"p_name_{i}") for i in range(num_players)]
            available_players_for_order = [name for name in player_name_inputs if name]
            st.divider()
            st.write("**Startreihenfolge für Spiel 1:**")
            initial_order_selection = st.multiselect("Wähle die Spieler in der gewünschten Zugreihenfolge für das erste Spiel aus:", options=available_players_for_order, default=[], key="initial_order_select", max_selections=num_players)
            submitted = st.form_submit_button("Turnier starten")
            if submitted:
                unique_names = set(filter(None, player_name_inputs))
                if len(unique_names) != num_players: st.error(f"Bitte genau {num_players} eindeutige Spielernamen eingeben.")
                elif len(initial_order_selection) != num_players: st.error(f"Bitte genau {num_players} Spieler für die Startreihenfolge auswählen.")
                elif len(set(initial_order_selection)) != num_players: st.error("Jeder Spieler darf in der Startreihenfolge nur einmal vorkommen.")
                else:
                    st.session_state.num_players = num_players
                    st.session_state.players = [{'id': i, 'name': name} for i, name in enumerate(player_name_inputs)]
                    recalculate_player_stats_from_games(st.session_state.players, [])
                    st.session_state.initial_turn_order = initial_order_selection
                    st.session_state.initialized = True
                    st.rerun()

    with col_import:
        st.subheader("Oder bestehendes Turnier importieren")
        uploaded_file = st.file_uploader("Wähle eine zuvor exportierte Excel-Datei aus.", type=['xlsx'])
        if uploaded_file is not None:
            if st.button("Turnier importieren"):
                with st.spinner("Importiere Turnierdaten..."):
                    if import_from_excel(uploaded_file):
                        st.success("Turnier erfolgreich importiert!")
                        import time
                        time.sleep(1)
                        st.rerun()

# --- App Hauptteil (Nach Spieler-Setup) ---
if st.session_state.initialized:
    current_game_number = len(st.session_state.get('games', [])) + 1
    if current_game_number == 1:
        turn_order_to_display = st.session_state.initial_turn_order
    else:
        if not st.session_state.next_turn_order_names and st.session_state.players:
            st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)
        turn_order_to_display = st.session_state.next_turn_order_names

    # --- NEUES LAYOUT: Spaltenaufteilung ---
    col1, col2 = st.columns([2, 3]) # Linke Spalte etwas breiter für das Formular

    # --- Linke Spalte: Aktion & Verlauf ---
    with col1:
        # ----- Spiel protokollieren -----
        if st.session_state.tournament_finished:
            st.subheader("🏁 Turnier beendet!")
            st.success("Das Turnier wurde manuell abgeschlossen. Es können keine weiteren Spiele protokolliert werden.")
        elif not turn_order_to_display:
            st.warning("Zugreihenfolge konnte nicht bestimmt werden.")
        else:
            st.subheader(f"📜 Spiel {current_game_number} protokollieren")
            with st.form("game_log_form"):
                st.write("**Zugreihenfolge für dieses Spiel:**")
                st.write(" → ".join(turn_order_to_display))
                
                selected_map_index = st.session_state.get('simulated_map_index', 0) if st.session_state.simulation_triggered else 0
                selected_map = st.selectbox("Karte auswählen", MAPS, index=selected_map_index, key=f"map_{current_game_number}")

                game_results_input = []
                for i, player_name in enumerate(turn_order_to_display):
                    st.markdown(f"**{i+1}. {player_name}**")
                    input_cols = st.columns(2)
                    with input_cols[0]:
                        default_faction_index = st.session_state.get('simulated_factions', {}).get(player_name, 0) if st.session_state.simulation_triggered else 0
                        selected_faction = st.selectbox(f"Fraktion für {player_name}", FACTIONS, index=default_faction_index, key=f"faction_{current_game_number}_{player_name}")
                    with input_cols[1]:
                        default_vp = st.session_state.get('simulated_vps', {}).get(player_name, 0) if st.session_state.simulation_triggered else 0
                        vp = st.number_input(f"Siegpunkte (VP) für {player_name}", min_value=0, step=1, value=default_vp, key=f"vp_{current_game_number}_{player_name}")
                    game_results_input.append({'name': player_name, 'faction': selected_faction, 'vp': vp})
                
                form_buttons_cols = st.columns(2)
                with form_buttons_cols[0]: simulate_button = st.form_submit_button("🎲 Spiel simulieren")
                with form_buttons_cols[1]: log_game_button = st.form_submit_button("💾 Spiel speichern", type="primary")

                if simulate_button:
                    st.session_state.simulation_triggered = True
                    st.session_state.simulated_map_index = random.choice(range(len(MAPS)))
                    available_factions = FACTIONS[:]
                    random.shuffle(available_factions)
                    if len(available_factions) >= len(turn_order_to_display):
                        st.session_state.simulated_factions = {player_name: FACTIONS.index(available_factions[i]) for i, player_name in enumerate(turn_order_to_display)}
                    else: st.warning("Nicht genügend Fraktionen für Simulation.")
                    if turn_order_to_display:
                        winner_name = random.choice(turn_order_to_display)
                        st.session_state.simulated_vps = {p: (30 if p == winner_name else random.randint(10, 29)) for p in turn_order_to_display}
                    st.rerun()

                if log_game_button:
                    st.session_state.simulation_triggered = False
                    
                    winners = [res['name'] for res in game_results_input if res['vp'] >= 30]
                    selected_factions = [res['faction'] for res in game_results_input]
                    faction_counts = Counter(selected_factions)
                    duplicates = [faction for faction, count in faction_counts.items() if count > 1]
                    
                    if len(winners) > 1: st.error(f"Fehler: Mehrere Spieler ({', '.join(winners)}) haben >= 30 VP.")
                    elif duplicates: st.error(f"Fehler: Fraktion(en) **{', '.join(duplicates)}** wurde(n) mehrfach ausgewählt.")
                    else:
                        sorted_results = sorted(game_results_input, key=lambda x: x['vp'], reverse=True)
                        game_log_entry = {'game_number': current_game_number, 'map': selected_map, 'turn_order': turn_order_to_display, 'results': []}
                        current_num_players = len(turn_order_to_display)
                        current_points_map = {rank: TOURNAMENT_POINTS_MAP.get(rank, 0) for rank in range(1, current_num_players + 1)}
                        
                        for rank, result in enumerate(sorted_results, 1):
                            result['rank'], result['tp'] = rank, current_points_map.get(rank, 0)
                            game_log_entry['results'].append(result)
                        
                        st.session_state.games.append(game_log_entry)
                        recalculate_player_stats_from_games(st.session_state.players, st.session_state.games)
                        st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)
                        st.rerun()
        
        st.divider()
        # ----- Spielprotokoll (Verlauf) -----
        st.subheader("📖 Spielprotokoll (Verlauf)")
        if not st.session_state.get('games', []): st.info("Noch keine Spiele protokolliert.")
        else:
            for game in reversed(st.session_state.get('games', [])):
                with st.expander(f"Spiel {game.get('game_number', 'N/A')} - Karte: {game.get('map', 'N/A')} (Zugreihenfolge: {' → '.join(game.get('turn_order', []))})"):
                    if game.get('results', []):
                        results_df = pd.DataFrame(game['results'])[['rank', 'name', 'faction', 'vp', 'tp']]
                        results_df.columns = ['Platz', 'Spieler', 'Fraktion', 'Siegpunkte (VP)', 'Turnierpunkte (TP)']
                        st.dataframe(results_df, hide_index=True, use_container_width=True)
                    else: st.write("Keine Ergebnisdaten vorhanden.")

    # --- Rechte Spalte: Dashboard & Analyse ---
    with col2:
        # ----- Rangliste -----
        st.subheader("📊 Aktuelle Rangliste")
        if st.session_state.players:
            standings_df = generate_standings_df(st.session_state.players)
            st.dataframe(standings_df, use_container_width=True, hide_index=True)
        else: st.warning("Keine Spielerdaten vorhanden.")
        
        st.divider()
        # ----- Punkteentwicklung -----
        st.subheader("📈 Punkteentwicklung")
        if not st.session_state.get('games', []): st.info("Noch keine Spiele gespielt.")
        elif not st.session_state.get('players', []): st.warning("Keine Spielerdaten vorhanden.")
        else:
            plot_df = generate_plot_data(st.session_state.games, st.session_state.players)
            if not plot_df.empty:
                fig = px.line(plot_df, x='Spiel', y='Kumulierte Turnierpunkte', color='Spieler', markers=True, title="Turnierpunkte über die Zeit", labels={'Spiel': 'Nach Spiel Nr.', 'Kumulierte Turnierpunkte': 'Turnierpunkte'})
                fig.update_layout(xaxis_tickmode='linear', title_font_size=16)
                st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        # ----- Detaillierte Statistiken -----
        if st.session_state.get('games', []):
            st.subheader("🔍 Detail-Statistiken")
            tab1, tab2 = st.tabs(["Fraktionen", "Karten"])
            with tab1:
                st.dataframe(calculate_faction_stats(st.session_state.games, FACTIONS), hide_index=True, use_container_width=True)
            with tab2:
                st.dataframe(calculate_map_stats(st.session_state.games, MAPS), hide_index=True, use_container_width=True)

    # --- Sidebar Optionen ---
    st.sidebar.title("Optionen")
    if st.session_state.initialized:
        if st.session_state.get('games', []):
            excel_data = {
                "Rangliste": generate_standings_df(st.session_state.players),
                "Fraktions-Statistik": calculate_faction_stats(st.session_state.games, FACTIONS),
                "Karten-Statistik": calculate_map_stats(st.session_state.games, MAPS),
                "Spielprotokolle": pd.DataFrame([
                    {"Spiel Nr": g['game_number'], "Karte": g['map'], "Spieler": r['name'], "Fraktion": r['faction'], "Siegpunkte (VP)": r['vp'], "Platz": r['rank'], "Turnierpunkte (TP)": r['tp'], "Zugreihenfolge (Spiel)": ' → '.join(g['turn_order'])}
                    for g in st.session_state.games for r in g['results']
                ])
            }
            st.sidebar.download_button(
                label="💾 Turnierdaten als Excel exportieren", data=df_to_excel(excel_data),
                file_name="root_turnier_ergebnisse.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else: st.sidebar.info("Noch keine Spiele gespielt zum Exportieren.")

        if not st.session_state.tournament_finished:
            st.sidebar.divider()
            if st.sidebar.button("🏁 Turnier manuell beenden", type="primary"):
                st.session_state.tournament_finished = True
                st.rerun()
        else: st.sidebar.success("Turnier wurde beendet.")