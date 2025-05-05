import streamlit as st
import pandas as pd
import openpyxl # Keep for Excel export even if not directly used in main logic
import plotly.express as px
import copy
from io import BytesIO
from collections import defaultdict, Counter # Counter hinzugefügt für Duplikat-Prüfung

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
    # Ggf. weitere Fraktionen aus Erweiterungen hinzufügen, falls gespielt
    # "Herr der Scharen", # Beispiel Marodeur
    # "Eisenwächter"     # Beispiel Marodeur
]
MAPS = ["Herbst", "Winter", "Berg", "See"] # Standard deutsche Kartenbezeichnungen
# Angepasste Punkteverteilung beibehalten (Hinweis: Ggf. für andere Spielerzahlen anpassen)
TOURNAMENT_POINTS_MAP = {1: 5, 2: 4, 3: 3, 4: 2, 5: 1}
# NUM_PLAYERS = 5 # Nicht mehr als globale Konstante benötigt

# --- Hilfsfunktionen ---
def initialize_state():
    """Initialisiert den Session State, falls noch nicht geschehen."""
    if 'players' not in st.session_state:
        st.session_state.players = [] # Liste von Dictionaries
    if 'games' not in st.session_state:
        st.session_state.games = [] # Liste von Game-Log Dictionaries
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
    # NEU: Zustände für Spieler- und Spielanzahl
    if 'num_players_input' not in st.session_state:
        st.session_state.num_players_input = 2 # Standardwert oder Minimum
    if 'total_games_input' not in st.session_state:
        st.session_state.total_games_input = 1 # Standardwert oder Minimum

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
    # Sortieren: Zuerst nach Turnierpunkten (aufsteigend), dann nach VP letztes Spiel (absteigend)
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
                'Fraktion': faction,
                'Gespielt': count,
                'Siege': data['wins'],
                'Ø Siegpunkte': f"{avg_vp:.2f}",
                'Ø Turnierpunkte': f"{avg_tp:.2f}"
            })
        else:
             stats_list.append({
                'Fraktion': faction,
                'Gespielt': 0,
                'Siege': 0,
                'Ø Siegpunkte': '-',
                'Ø Turnierpunkte': '-'
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
                'Karte': map_name,
                'Gespielt (Spiele)': data['count'],
                'Ø Siegpunkte (Gesamt)': f"{avg_vp:.2f}"
            })
        else:
             stats_list.append({
                'Karte': map_name,
                'Gespielt (Spiele)': 0,
                'Ø Siegpunkte (Gesamt)': '-'
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

# Initialisierung des States
initialize_state()

# --- 1. Spieler-Setup (Nur am Anfang) ---
if not st.session_state.initialized:
    st.subheader("Turnier Setup")
    with st.form("player_form"):
        # Eingabe für Spieler- und Spielanzahl
        num_players = st.number_input("Anzahl der Spieler", min_value=2, max_value=10, value=st.session_state.num_players_input, step=1, key="num_players_selector")
        total_games = st.number_input("Gesamtzahl der Spiele", min_value=1, value=st.session_state.total_games_input, step=1, key="total_games_selector")

        # Update state if changed
        st.session_state.num_players_input = num_players
        st.session_state.total_games_input = total_games

        st.divider()
        st.write("**Spielernamen eingeben:**")
        player_name_inputs = {}
        temp_player_list = [""] * num_players # Dynamische Größe
        for i in range(num_players):
            player_name_inputs[i] = st.text_input(f"Name Spieler {i+1}", key=f"p_name_{i}")
            if player_name_inputs[i]:
                temp_player_list[i] = player_name_inputs[i]

        available_players_for_order = [name for name in temp_player_list if name]

        st.divider()
        st.write("**Startreihenfolge für Spiel 1:**")
        initial_order_selection = st.multiselect(
            "Wähle die Spieler in der gewünschten Zugreihenfolge für das erste Spiel aus:",
            options=available_players_for_order,
            default=[],
            key="initial_order_select",
            max_selections=num_players # Dynamische max_selections
        )

        submitted = st.form_submit_button("Turnier starten")
        if submitted:
            final_player_names = list(player_name_inputs.values())
            unique_names = set(filter(None, final_player_names))
            # Validierung gegen die ausgewählte Spieleranzahl
            if len(unique_names) != num_players:
                st.error(f"Bitte genau {num_players} eindeutige Spielernamen eingeben.")
            elif len(initial_order_selection) != num_players:
                 st.error(f"Bitte genau {num_players} Spieler für die Startreihenfolge auswählen.")
            elif len(set(initial_order_selection)) != num_players:
                 st.error("Jeder Spieler darf in der Startreihenfolge nur einmal vorkommen.")
            else:
                # Speichere die finale Anzahl Spieler und Spiele
                st.session_state.num_players = num_players
                st.session_state.total_games = total_games
                # Initialisiere Spielerdaten
                st.session_state.players = [
                    {
                        'id': i,
                        'name': name,
                        'total_tp': 0, 'total_vp': 0, 'wins': 0, 'last_vp': 0,
                        'played_factions': [], 'played_factions_str': '',
                        'total_placement_sum': 0, 'games_played': 0
                     } for i, name in enumerate(final_player_names)
                ]
                st.session_state.initial_turn_order = initial_order_selection
                st.session_state.next_turn_order_names = []
                st.session_state.initialized = True
                st.rerun()

# --- App Hauptteil (Nach Spieler-Setup) ---
if st.session_state.initialized:

    current_game_number = len(st.session_state.get('games', [])) + 1

    # Bestimme die anzuzeigende Zugreihenfolge
    if current_game_number == 1:
        turn_order_to_display = st.session_state.initial_turn_order
    else:
        # Berechne nur, wenn nötig und Daten vorhanden
        if not st.session_state.next_turn_order_names and st.session_state.players:
             st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)
        turn_order_to_display = st.session_state.next_turn_order_names

    # Define layout columns
    col1, col2 = st.columns([2, 3]) # Verhältnis ggf. anpassen

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
        # Prüfen, ob das Turnier beendet ist
        if current_game_number > st.session_state.get('total_games', 0):
            st.subheader("🏁 Turnier beendet!")
            st.success(f"Alle {st.session_state.total_games} Spiele wurden protokolliert.")
            # Optional: Endrangliste nochmal anzeigen?
            if 'players' in st.session_state and st.session_state.players:
                st.write("Abschlusstabelle:")
                final_standings_df = generate_standings_df(st.session_state.players)
                st.dataframe(final_standings_df, use_container_width=True, hide_index=True)

        elif not turn_order_to_display:
             st.warning("Zugreihenfolge konnte nicht bestimmt werden.")
        else:
            # Zeige das Formular zum Protokollieren des nächsten Spiels
            st.subheader(f"📜 Spiel {current_game_number} von {st.session_state.total_games} protokollieren")

            if st.session_state.show_faction_warning:
                 for msg in st.session_state.warning_messages:
                     st.warning(msg)
                 st.session_state.show_faction_warning = False
                 st.session_state.warning_messages = []

            with st.form("game_log_form"):
                st.write("**Zugreihenfolge für dieses Spiel:**")
                st.write(" → ".join(turn_order_to_display))

                selected_map = st.selectbox("Karte auswählen", MAPS, key=f"map_{current_game_number-1}")

                game_results_input = []
                faction_warning_check = []

                for i, player_name in enumerate(turn_order_to_display):
                    st.markdown(f"**{i+1}. {player_name}**")
                    input_cols = st.columns(2)
                    with input_cols[0]:
                        player_data = get_player_data_by_name(player_name)
                        played_before = player_data.get('played_factions', []) if player_data else []
                        selected_faction = st.selectbox(f"Fraktion für {player_name}", FACTIONS, key=f"faction_{current_game_number-1}_{player_name}")
                        if selected_faction in played_before:
                             faction_warning_check.append(f"Spieler **{player_name}** hat die Fraktion **{selected_faction}** bereits gespielt!")

                    with input_cols[1]:
                         vp = st.number_input(f"Siegpunkte (VP) für {player_name}", min_value=0, step=1, key=f"vp_{current_game_number-1}_{player_name}")

                    game_results_input.append({'name': player_name, 'faction': selected_faction, 'vp': vp})

                log_game_button = st.form_submit_button("Spiel speichern")

                if log_game_button:
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
                                'game_number': current_game_number,
                                'map': selected_map,
                                'turn_order': copy.deepcopy(turn_order_to_display),
                                'results': []
                            }

                            for rank, result in enumerate(sorted_results, 1):
                                player_name = result.get('name')
                                if not player_name: continue
                                player_data = get_player_data_by_name(player_name)
                                if player_data:
                                    # Verwende .get(rank, 0) um sicherzustellen, dass auch bei >5 Spielern ein Wert (0) zurückkommt
                                    tp = TOURNAMENT_POINTS_MAP.get(rank, 0)
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
                            # Berechne die nächste Zugreihenfolge NACH dem Speichern
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

    # --- Export Funktion ---
    st.sidebar.title("Export")
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
                     "Spiel Nr": game_num,
                     "Karte": game_map,
                     "Spieler": result.get('name', 'N/A'),
                     "Fraktion": result.get('faction', 'N/A'),
                     "Siegpunkte (VP)": result.get('vp', 0),
                     "Platz": result.get('rank', 0),
                     "Turnierpunkte (TP)": result.get('tp', 0),
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

