import streamlit as st
import pandas as pd
import openpyxl # Keep for Excel export
import plotly.express as px
import copy
import random # Hinzugefügt für Simulation
from io import BytesIO
from collections import defaultdict, Counter

# --- Konstanten (Deutsche Bezeichnungen) ---
FACTIONS = [
    "Marquise de Katz",
    "Baumkronen-Dynastie",
    "Waldland-Allianz",
    "Vagabund",
    "Flussvolk-Kompanie",
    "Echsen-Kult",
    "Untergrund-Herzogtum",
    "Krähen-Komplott",
    # Ggf. weitere Fraktionen aus Erweiterungen hinzufügen
]
MAPS = ["Herbst", "Winter", "Berg", "See"]
# Angepasste Punkteverteilung
TOURNAMENT_POINTS_MAP = {1: 5, 2: 4, 3: 3, 4: 2, 5: 1}

# --- Hilfsfunktionen ---
def initialize_state():
    """Initialisiert den Session State, falls noch nicht geschehen."""
    if 'players' not in st.session_state:
        st.session_state.players = []
    if 'games' not in st.session_state:
        st.session_state.games = []
    if 'next_turn_order_names' not in st.session_state:
        st.session_state.next_turn_order_names = []
    if 'initial_turn_order' not in st.session_state:
        st.session_state.initial_turn_order = []
    if 'initialized' not in st.session_state:
        st.session_state.initialized = False
    if 'show_faction_warning' not in st.session_state:
        st.session_state.show_faction_warning = False
    if 'warning_messages' not in st.session_state:
         st.session_state.warning_messages = []
    if 'num_players_input' not in st.session_state:
        st.session_state.num_players_input = 2
    if 'tournament_finished' not in st.session_state:
        st.session_state.tournament_finished = False
    # NEU: Zustände für simulierte Werte
    if 'simulated_map_index' not in st.session_state:
        st.session_state.simulated_map_index = 0 # Standardmäßig erste Karte
    if 'simulated_factions' not in st.session_state:
        st.session_state.simulated_factions = {} # Dict: {player_name: faction_index}
    if 'simulated_vps' not in st.session_state:
        st.session_state.simulated_vps = {} # Dict: {player_name: vp}
    if 'simulation_triggered' not in st.session_state:
        st.session_state.simulation_triggered = False # Flag, um zu wissen, ob simuliert wurde


def get_player_data_by_name(name):
    """Gibt die Daten eines Spielers anhand des Namens zurück."""
    for player in st.session_state.players:
        if player['name'] == name:
            return player
    return None

def calculate_next_turn_order(players):
    """Berechnet die Zugreihenfolge für das nächste Spiel (ab Spiel 2)."""
    if not players:
        return []
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
        for p in display_players:
            p['avg_placement'] = '-'

    df = pd.DataFrame(display_players)

    expected_cols = ['name', 'total_tp', 'total_vp', 'wins', 'last_vp', 'avg_placement', 'played_factions_str']
    if not all(col in df.columns for col in expected_cols):
        st.error("Fehler bei der DataFrame-Erstellung. Nicht alle erwarteten Spalten sind vorhanden.")
        return pd.DataFrame(columns=['Rang', 'Name', 'Ges. Turnierpkt.', 'Ges. Siegpunkte', 'Siege', 'Letzte Spiel VP', 'Ø Platzierung', 'Gespielte Fraktionen'])

    df = df.sort_values(by='total_tp', ascending=False).reset_index(drop=True)
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

    for i, game in enumerate(games):
        temp_player_points = {name: player_points_over_time[name][-1] for name in player_names}
        for player_result in game.get('results', []):
            player_name = player_result.get('name')
            tp = player_result.get('tp', 0)
            if player_name in temp_player_points:
                 temp_player_points[player_name] += tp
        for player_name, total_points in temp_player_points.items():
             if player_name in player_points_over_time:
                player_points_over_time[player_name].append(total_points)

    for player_name, points_list in player_points_over_time.items():
        for game_idx, points in enumerate(points_list):
            plot_data.append({'Spiel': game_idx, 'Spieler': player_name, 'Kumulierte Turnierpunkte': points})

    if not plot_data: return pd.DataFrame()
    return pd.DataFrame(plot_data)

def calculate_faction_stats(games, factions):
    """Berechnet Statistiken für jede Fraktion."""
    if not games:
        return pd.DataFrame(columns=['Fraktion', 'Gespielt', 'Siege', 'Ø Siegpunkte', 'Ø Turnierpunkte'])

    faction_data = defaultdict(lambda: {'count': 0, 'total_vp': 0, 'total_tp': 0, 'wins': 0})

    for game in games:
        for result in game.get('results', []):
            faction = result.get('faction')
            vp = result.get('vp', 0)
            tp = result.get('tp', 0)
            rank = result.get('rank')
            if faction:
                faction_data[faction]['count'] += 1
                faction_data[faction]['total_vp'] += vp
                faction_data[faction]['total_tp'] += tp
                if rank == 1:
                    faction_data[faction]['wins'] += 1

    stats_list = []
    for faction in factions:
        data = faction_data[faction]
        count = data['count']
        if count > 0:
            avg_vp = data['total_vp'] / count
            avg_tp = data['total_tp'] / count
            stats_list.append({
                'Fraktion': faction, 'Gespielt': count, 'Siege': data['wins'],
                'Ø Siegpunkte': f"{avg_vp:.2f}", 'Ø Turnierpunkte': f"{avg_tp:.2f}"
            })
        else:
             stats_list.append({
                'Fraktion': faction, 'Gespielt': 0, 'Siege': 0,
                'Ø Siegpunkte': '-', 'Ø Turnierpunkte': '-'
            })

    df = pd.DataFrame(stats_list)
    df = df.sort_values(by='Gespielt', ascending=False).reset_index(drop=True)
    return df

def calculate_map_stats(games, maps):
    """Berechnet Statistiken für jede Karte."""
    if not games:
        return pd.DataFrame(columns=['Karte', 'Gespielt', 'Ø Siegpunkte (Gesamt)'])

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
            stats_list.append({
                'Karte': map_name, 'Gespielt (Spiele)': data['count'],
                'Ø Siegpunkte (Gesamt)': f"{avg_vp:.2f}"
            })
        else:
             stats_list.append({
                'Karte': map_name, 'Gespielt (Spiele)': 0, 'Ø Siegpunkte (Gesamt)': '-'
            })

    df = pd.DataFrame(stats_list)
    df = df.sort_values(by='Gespielt (Spiele)', ascending=False).reset_index(drop=True)
    return df

def df_to_excel(df_dict):
    """Exportiert mehrere DataFrames in eine Excel-Datei mit mehreren Blättern."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                 df.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    return processed_data

# --- Streamlit App ---
st.set_page_config(page_title="Root Turnier Manager", layout="wide", initial_sidebar_state="expanded")
st.title("🏆 Root Brettspiel Turnier Manager")

initialize_state()

# --- 1. Spieler-Setup (Nur am Anfang) ---
if not st.session_state.initialized:
    st.subheader("Turnier Setup")
    with st.form("player_form"):
        num_players = st.number_input("Anzahl der Spieler", min_value=2, max_value=10, value=st.session_state.num_players_input, step=1, key="num_players_selector")
        st.session_state.num_players_input = num_players

        st.divider()
        st.write("**Spielernamen eingeben:**")
        player_name_inputs = {}
        temp_player_list = [""] * num_players
        for i in range(num_players):
            player_name_inputs[i] = st.text_input(f"Name Spieler {i+1}", key=f"p_name_{i}")
            if player_name_inputs[i]:
                temp_player_list[i] = player_name_inputs[i]

        available_players_for_order = [name for name in temp_player_list if name]

        st.divider()
        st.write("**Startreihenfolge für Spiel 1:**")
        initial_order_selection = st.multiselect(
            "Wähle die Spieler in der gewünschten Zugreihenfolge für das erste Spiel aus:",
            options=available_players_for_order, default=[],
            key="initial_order_select", max_selections=num_players
        )

        submitted = st.form_submit_button("Turnier starten")
        if submitted:
            final_player_names = list(player_name_inputs.values())
            unique_names = set(filter(None, final_player_names))
            if len(unique_names) != num_players:
                st.error(f"Bitte genau {num_players} eindeutige Spielernamen eingeben.")
            elif len(initial_order_selection) != num_players:
                 st.error(f"Bitte genau {num_players} Spieler für die Startreihenfolge auswählen.")
            elif len(set(initial_order_selection)) != num_players:
                 st.error("Jeder Spieler darf in der Startreihenfolge nur einmal vorkommen.")
            else:
                st.session_state.num_players = num_players
                st.session_state.players = [
                    {'id': i, 'name': name, 'total_tp': 0, 'total_vp': 0, 'wins': 0,
                     'last_vp': 0, 'played_factions': [], 'played_factions_str': '',
                     'total_placement_sum': 0, 'games_played': 0
                    } for i, name in enumerate(final_player_names)
                ]
                st.session_state.initial_turn_order = initial_order_selection
                st.session_state.next_turn_order_names = []
                st.session_state.tournament_finished = False
                st.session_state.initialized = True
                # Reset simulation state on new tournament start
                st.session_state.simulated_map_index = 0
                st.session_state.simulated_factions = {}
                st.session_state.simulated_vps = {}
                st.session_state.simulation_triggered = False
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

    col1, col2 = st.columns([2, 3])

    with col1:
        st.subheader("📊 Aktuelle Rangliste")
        if 'players' in st.session_state and st.session_state.players:
            standings_df = generate_standings_df(st.session_state.players)
            st.dataframe(standings_df, use_container_width=True, hide_index=True)
        else:
            st.warning("Keine Spielerdaten vorhanden für die Rangliste.")

        st.subheader("📈 Punkteentwicklung (Turnierpunkte)")
        if not st.session_state.get('games', []):
             st.info("Noch keine Spiele gespielt, um die Entwicklung anzuzeigen.")
        elif not st.session_state.get('players', []):
             st.warning("Keine Spielerdaten vorhanden für die Punkteentwicklung.")
        else:
            plot_df = generate_plot_data(st.session_state.games, st.session_state.players)
            if not plot_df.empty:
                fig = px.line(plot_df, x='Spiel', y='Kumulierte Turnierpunkte', color='Spieler',
                              markers=True, title="Turnierpunkte über die Zeit",
                              labels={'Spiel': 'Nach Spiel Nr.', 'Kumulierte Turnierpunkte': 'Turnierpunkte'})
                fig.update_layout(xaxis_tickmode = 'linear')
                st.plotly_chart(fig, use_container_width=True)

    with col2:
        if st.session_state.tournament_finished:
            st.subheader("🏁 Turnier beendet!")
            st.success("Das Turnier wurde manuell abgeschlossen.")
            if 'players' in st.session_state and st.session_state.players:
                st.write("Abschlusstabelle:")
                final_standings_df = generate_standings_df(st.session_state.players)
                st.dataframe(final_standings_df, use_container_width=True, hide_index=True)

        elif not turn_order_to_display:
             st.warning("Zugreihenfolge konnte nicht bestimmt werden.")
        else:
            st.subheader(f"📜 Spiel {current_game_number} protokollieren")

            if st.session_state.show_faction_warning:
                 for msg in st.session_state.warning_messages:
                     st.warning(msg)
                 st.session_state.show_faction_warning = False
                 st.session_state.warning_messages = []

            # Form für Spieleingabe
            with st.form("game_log_form"):
                st.write("**Zugreihenfolge für dieses Spiel:**")
                st.write(" → ".join(turn_order_to_display))

                # Karte auswählen (mit Vorbelegung durch Simulation)
                map_key = f"map_{current_game_number-1}"
                selected_map_index = st.session_state.get('simulated_map_index', 0) if st.session_state.simulation_triggered else 0
                selected_map = st.selectbox("Karte auswählen", MAPS, key=map_key, index=selected_map_index)

                game_results_input = []
                faction_warning_check = []
                num_players_in_turn_order = len(turn_order_to_display)
                current_points_map = {rank: TOURNAMENT_POINTS_MAP.get(rank, 0) for rank in range(1, num_players_in_turn_order + 1)}

                # Eingabefelder für jeden Spieler
                for i, player_name in enumerate(turn_order_to_display):
                    st.markdown(f"**{i+1}. {player_name}**")
                    input_cols = st.columns(2)
                    with input_cols[0]:
                        player_data = get_player_data_by_name(player_name)
                        played_before = player_data.get('played_factions', []) if player_data else []
                        faction_key = f"faction_{current_game_number-1}_{player_name}"
                        # Fraktion auswählen (mit Vorbelegung durch Simulation)
                        default_faction_index = st.session_state.get('simulated_factions', {}).get(player_name, 0) if st.session_state.simulation_triggered else 0
                        selected_faction = st.selectbox(f"Fraktion für {player_name}", FACTIONS, key=faction_key, index=default_faction_index)
                        if selected_faction in played_before:
                             faction_warning_check.append(f"Spieler **{player_name}** hat die Fraktion **{selected_faction}** bereits gespielt!")

                    with input_cols[1]:
                         vp_key = f"vp_{current_game_number-1}_{player_name}"
                         # VP eingeben (mit Vorbelegung durch Simulation)
                         default_vp = st.session_state.get('simulated_vps', {}).get(player_name, 0) if st.session_state.simulation_triggered else 0
                         vp = st.number_input(f"Siegpunkte (VP) für {player_name}", min_value=0, step=1, key=vp_key, value=default_vp)

                    game_results_input.append({'name': player_name, 'faction': selected_faction, 'vp': vp})

                # Buttons am Ende des Formulars
                form_buttons_cols = st.columns(2)
                with form_buttons_cols[0]:
                    simulate_button = st.form_submit_button("🎲 Spiel simulieren")
                with form_buttons_cols[1]:
                    log_game_button = st.form_submit_button("💾 Spiel speichern", type="primary")


                # --- Simulationslogik ---
                if simulate_button:
                    st.session_state.simulation_triggered = True # Flag setzen

                    # 1. Karte simulieren
                    sim_map_index = random.choice(range(len(MAPS)))
                    st.session_state.simulated_map_index = sim_map_index

                    # 2. Fraktionen simulieren (einzigartig pro Spiel)
                    available_factions = FACTIONS[:] # Kopie erstellen
                    random.shuffle(available_factions)
                    sim_factions = {}
                    if len(available_factions) >= len(turn_order_to_display):
                        for i, player_name in enumerate(turn_order_to_display):
                            chosen_faction = available_factions[i]
                            # Finde den Index der gewählten Fraktion in der Original-FACTIONS-Liste
                            sim_factions[player_name] = FACTIONS.index(chosen_faction)
                        st.session_state.simulated_factions = sim_factions
                    else:
                         st.warning("Nicht genügend einzigartige Fraktionen für die Simulation verfügbar.")
                         st.session_state.simulated_factions = {} # Reset

                    # 3. Siegpunkte simulieren (1 Gewinner mit 30, Rest 10-29)
                    sim_vps = {}
                    if turn_order_to_display: # Nur wenn Spieler vorhanden
                        winner_name = random.choice(turn_order_to_display)
                        for player_name in turn_order_to_display:
                            if player_name == winner_name:
                                sim_vps[player_name] = 30
                            else:
                                sim_vps[player_name] = random.randint(10, 29)
                        st.session_state.simulated_vps = sim_vps
                    else:
                        st.session_state.simulated_vps = {} # Reset

                    # Wichtig: Nach dem Setzen der Simulationswerte im State, muss Streamlit neu laden,
                    # damit die Widgets die neuen Default-Werte anzeigen.
                    st.rerun()


                # --- Speicherlogik ---
                if log_game_button:
                    # Setze Simulationsflag zurück, damit beim nächsten Laden keine Vorbelegung erfolgt
                    st.session_state.simulation_triggered = False
                    st.session_state.simulated_map_index = 0
                    st.session_state.simulated_factions = {}
                    st.session_state.simulated_vps = {}

                    # --- VALIDIERUNGEN ---
                    winners = [res['name'] for res in game_results_input if res['vp'] >= 30]
                    if len(winners) > 1:
                        st.error(f"Fehler: Mehr als ein Spieler ({', '.join(winners)}) hat 30 oder mehr Siegpunkte erreicht. Nur ein Spieler kann das Spiel auf diese Weise gewinnen. Bitte korrigieren.")
                    else:
                        selected_factions_this_game = [result['faction'] for result in game_results_input]
                        if len(selected_factions_this_game) != len(set(selected_factions_this_game)):
                            counts = Counter(selected_factions_this_game)
                            duplicates = [faction for faction, count in counts.items() if count > 1]
                            st.error(f"Fehler: Die Fraktion(en) **{', '.join(duplicates)}** wurde(n) mehrfach ausgewählt. Jede Fraktion darf pro Spiel nur einmal vorkommen. Bitte korrigieren.")
                        else:
                            # --- Nur fortfahren, wenn beide Checks OK sind ---
                            if faction_warning_check:
                                st.session_state.show_faction_warning = True
                                st.session_state.warning_messages = faction_warning_check

                            sorted_results = sorted(game_results_input, key=lambda x: x['vp'], reverse=True)

                            game_log_entry = {
                                'game_number': current_game_number, 'map': selected_map,
                                'turn_order': copy.deepcopy(turn_order_to_display), 'results': []
                            }

                            for rank, result in enumerate(sorted_results, 1):
                                player_name = result.get('name')
                                if not player_name: continue
                                player_data = get_player_data_by_name(player_name)
                                if player_data:
                                    tp = current_points_map.get(rank, 0)
                                    result['rank'] = rank
                                    result['tp'] = tp
                                    game_log_entry['results'].append(result)

                                    player_data['total_tp'] = player_data.get('total_tp', 0) + tp
                                    player_data['total_vp'] = player_data.get('total_vp', 0) + result.get('vp', 0)
                                    player_data['last_vp'] = result.get('vp', 0)
                                    current_faction = result.get('faction')
                                    if current_faction and current_faction not in player_data.get('played_factions', []):
                                         player_data.setdefault('played_factions', []).append(current_faction)
                                    player_data['played_factions_str'] = ", ".join(player_data.get('played_factions', []))
                                    player_data['total_placement_sum'] = player_data.get('total_placement_sum', 0) + rank
                                    player_data['games_played'] = player_data.get('games_played', 0) + 1
                                    if rank == 1:
                                        player_data['wins'] = player_data.get('wins', 0) + 1

                            st.session_state.setdefault('games', []).append(game_log_entry)
                            st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)
                            st.rerun()

            # --- Zusätzliche Auswertungen (nur anzeigen, wenn Spiele vorhanden) ---
            if st.session_state.get('games', []):
                st.subheader("📊 Fraktions-Statistiken")
                faction_stats_df = calculate_faction_stats(st.session_state.games, FACTIONS)
                st.dataframe(faction_stats_df, hide_index=True, use_container_width=True)

                st.subheader("🗺️ Karten-Statistiken")
                map_stats_df = calculate_map_stats(st.session_state.games, MAPS)
                st.dataframe(map_stats_df, hide_index=True, use_container_width=True)

            # --- Spielprotokoll anzeigen ---
            st.subheader("📖 Spielprotokoll")
            if not st.session_state.get('games', []):
                st.info("Noch keine Spiele protokolliert.")
            else:
                for game in reversed(st.session_state.get('games', [])):
                    game_num = game.get('game_number', 'N/A')
                    game_map = game.get('map', 'N/A')
                    turn_order = game.get('turn_order', [])
                    expander_title = (f"Spiel {game_num} - Karte: {game_map} "
                                      f"(Zugreihenfolge: {' → '.join(turn_order)})")
                    with st.expander(expander_title):
                        results_data = game.get('results', [])
                        if results_data:
                            results_df = pd.DataFrame(results_data)
                            display_cols = {
                                'rank': 'Platz', 'name': 'Spieler', 'faction': 'Fraktion',
                                'vp': 'Siegpunkte (VP)', 'tp': 'Turnierpunkte (TP)'
                            }
                            cols_to_display = [col for col in display_cols.keys() if col in results_df.columns]
                            results_df_display = results_df[cols_to_display].rename(columns=display_cols)
                            st.dataframe(results_df_display, hide_index=True, use_container_width=True)
                        else:
                            st.write("Keine Ergebnisdaten für dieses Spiel vorhanden.")

    # --- Sidebar mit Export und Beenden-Button ---
    st.sidebar.title("Optionen")

    # Export Funktion
    if st.session_state.initialized and st.session_state.get('games', []):
        excel_data = {}
        if st.session_state.get('players', []):
            standings_export_df = generate_standings_df(st.session_state.players)
            excel_data["Rangliste"] = standings_export_df
        else:
             excel_data["Rangliste"] = None

        faction_stats_export_df = calculate_faction_stats(st.session_state.games, FACTIONS)
        excel_data["Fraktions-Statistik"] = faction_stats_export_df

        map_stats_export_df = calculate_map_stats(st.session_state.games, MAPS)
        excel_data["Karten-Statistik"] = map_stats_export_df

        all_games_list = []
        for game in st.session_state.get('games', []):
             turn_order_str = ", ".join(game.get('turn_order', []))
             game_num = game.get('game_number', 'N/A')
             game_map = game.get('map', 'N/A')
             for result in game.get('results', []):
                 all_games_list.append({
                     "Spiel Nr": game_num, "Karte": game_map, "Spieler": result.get('name', 'N/A'),
                     "Fraktion": result.get('faction', 'N/A'), "Siegpunkte (VP)": result.get('vp', 0),
                     "Platz": result.get('rank', 0), "Turnierpunkte (TP)": result.get('tp', 0),
                     "Zugreihenfolge (Spiel)": turn_order_str
                 })
        if all_games_list:
            excel_data["Spielprotokolle"] = pd.DataFrame(all_games_list)
        else:
            excel_data["Spielprotokolle"] = None

        excel_bytes = df_to_excel(excel_data)

        st.sidebar.download_button(
            label="💾 Turnierdaten als Excel exportieren",
            data=excel_bytes,
            file_name="root_turnier_ergebnisse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    elif st.session_state.initialized:
        st.sidebar.info("Noch keine Spiele gespielt zum Exportieren.")

    # Button zum Beenden des Turniers
    if st.session_state.initialized and not st.session_state.tournament_finished:
        st.sidebar.divider()
        if st.sidebar.button("🏁 Turnier manuell beenden", type="primary"):
            st.session_state.tournament_finished = True
            st.rerun()
    elif st.session_state.tournament_finished:
         st.sidebar.success("Turnier wurde beendet.")

