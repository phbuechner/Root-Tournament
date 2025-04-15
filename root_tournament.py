import streamlit as st
import pandas as pd
import plotly.express as px
import copy
from io import BytesIO
from collections import defaultdict, Counter # Counter hinzugefügt für Duplikat-Prüfung

# --- Konstanten ---
FACTIONS = ["Marquise de Cat", "Eyrie Dynasties", "Woodland Alliance", "Vagabond", "Riverfolk Company", "Lizard Cult", "Underground Duchy", "Corvid Conspiracy"]
MAPS = ["Autumn", "Winter", "Mountain", "Lake"] # Ggf. anpassen/eindeutschen
TOURNAMENT_POINTS_MAP = {1: 10, 2: 7, 3: 5, 4: 3, 5: 1}
NUM_PLAYERS = 5

# --- Hilfsfunktionen ---
def initialize_state():
    """Initialisiert den Session State, falls noch nicht geschehen."""
    if 'players' not in st.session_state:
        st.session_state.players = [] # Liste von Dictionaries
    if 'games' not in st.session_state:
        st.session_state.games = [] # Liste von Game-Log Dictionaries
    if 'next_turn_order_names' not in st.session_state:
        st.session_state.next_turn_order_names = []
    if 'initialized' not in st.session_state:
        st.session_state.initialized = False
    if 'show_faction_warning' not in st.session_state:
        st.session_state.show_faction_warning = False
    if 'warning_messages' not in st.session_state:
         st.session_state.warning_messages = []

def get_player_data_by_name(name):
    """Gibt die Daten eines Spielers anhand des Namens zurück."""
    for player in st.session_state.players:
        if player['name'] == name:
            return player
    return None

def calculate_next_turn_order(players):
    """Berechnet die Zugreihenfolge für das nächste Spiel."""
    if not players:
        return []
    # Sortieren: Zuerst nach Turnierpunkten (aufsteigend), dann nach VP letztes Spiel (absteigend)
    sorted_players = sorted(players, key=lambda p: (p['total_tp'], -p.get('last_vp', 0)))
    return [p['name'] for p in sorted_players]

def generate_standings_df(players):
    """Erstellt ein Pandas DataFrame für die Rangliste."""
    if not players:
        return pd.DataFrame(columns=['Rang', 'Name', 'Ges. Turnierpkt.', 'Ges. Siegpunkte', 'Siege', 'Letzte Spiel VP', 'Ø Platzierung', 'Gespielte Fraktionen'])

    # Kopie erstellen, um Originaldaten nicht zu ändern
    display_players = copy.deepcopy(players)

    # Berechne Ø Platzierung
    num_games = len(st.session_state.games)
    if num_games > 0:
        for p in display_players:
             p['avg_placement'] = f"{p['total_placement_sum'] / p['games_played']:.2f}" if p['games_played'] > 0 else '-'
    else:
        for p in display_players:
            p['avg_placement'] = '-'


    df = pd.DataFrame(display_players)
    # Sortieren für die Anzeige
    df = df.sort_values(by='total_tp', ascending=False).reset_index(drop=True)
    df['Rang'] = df.index + 1
    # Spalten auswählen und umbenennen
    df = df[['Rang', 'name', 'total_tp', 'total_vp', 'wins', 'last_vp', 'avg_placement', 'played_factions_str']]
    df.columns = ['Rang', 'Name', 'Ges. Turnierpkt.', 'Ges. Siegpunkte', 'Siege', 'Letzte Spiel VP', 'Ø Platzierung', 'Gespielte Fraktionen']
    return df

# Die Funktion highlight_top3 wurde entfernt.

def generate_plot_data(games, players):
    """Bereitet Daten für das Plotly-Diagramm vor."""
    plot_data = []
    # Initialize player points over time, starting at 0 before game 1
    player_points_over_time = {p['name']: [0] for p in players}

    for i, game in enumerate(games):
        game_number = i + 1
        # Get the current points total for each player before this game
        temp_player_points = {p['name']: player_points_over_time[p['name']][-1] for p in players}

        # Add points earned in this game
        for player_result in game['results']:
            player_name = player_result['name']
            if player_name in temp_player_points: # Ensure player exists
                 temp_player_points[player_name] += player_result['tp']

        # Append the new total points after the game
        for player_name, total_points in temp_player_points.items():
             if player_name in player_points_over_time: # Ensure player exists
                player_points_over_time[player_name].append(total_points)

    # Convert data to long format suitable for Plotly
    for player_name, points_list in player_points_over_time.items():
        for game_idx, points in enumerate(points_list): # game_idx 0 is the starting point
            plot_data.append({'Spiel': game_idx, 'Spieler': player_name, 'Kumulierte Turnierpunkte': points})

    return pd.DataFrame(plot_data)

def calculate_faction_stats(games, factions):
    """Berechnet Statistiken für jede Fraktion."""
    if not games:
        return pd.DataFrame(columns=['Fraktion', 'Gespielt', 'Siege', 'Ø Siegpunkte', 'Ø Turnierpunkte'])

    # Use defaultdict for easier aggregation
    faction_data = defaultdict(lambda: {'count': 0, 'total_vp': 0, 'total_tp': 0, 'wins': 0})

    # Aggregate data from all games
    for game in games:
        for result in game['results']:
            faction = result['faction']
            faction_data[faction]['count'] += 1
            faction_data[faction]['total_vp'] += result['vp']
            faction_data[faction]['total_tp'] += result['tp']
            if result['rank'] == 1:
                faction_data[faction]['wins'] += 1

    # Prepare data for DataFrame
    stats_list = []
    for faction in factions: # Iterate through all possible factions
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
             # Include factions that were never played
             stats_list.append({
                'Fraktion': faction,
                'Gespielt': 0,
                'Siege': 0,
                'Ø Siegpunkte': '-',
                'Ø Turnierpunkte': '-'
            })


    df = pd.DataFrame(stats_list)
    # Sort by times played for relevance
    df = df.sort_values(by='Gespielt', ascending=False).reset_index(drop=True)
    return df

def calculate_map_stats(games, maps):
    """Berechnet Statistiken für jede Karte."""
    if not games:
        return pd.DataFrame(columns=['Karte', 'Gespielt', 'Ø Siegpunkte (Gesamt)'])

    # Use defaultdict for easier aggregation
    map_data = defaultdict(lambda: {'count': 0, 'total_vp': 0, 'player_games': 0}) # player_games tracks player entries on this map

    # Aggregate data from all games
    for game in games:
        game_map = game['map']
        map_data[game_map]['count'] += 1 # Counts how many games used this map
        for result in game['results']:
            map_data[game_map]['total_vp'] += result['vp']
            map_data[game_map]['player_games'] += 1 # Counts total player results on this map

    # Prepare data for DataFrame
    stats_list = []
    for map_name in maps: # Iterate through all possible maps
        data = map_data[map_name]
        player_games_on_map = data['player_games']
        if player_games_on_map > 0:
            avg_vp = data['total_vp'] / player_games_on_map # Average VP across all players in games on this map
            stats_list.append({
                'Karte': map_name,
                'Gespielt (Spiele)': data['count'], # How many games used this map
                'Ø Siegpunkte (Gesamt)': f"{avg_vp:.2f}"
            })
        else:
            # Include maps that were never played
             stats_list.append({
                'Karte': map_name,
                'Gespielt (Spiele)': 0,
                'Ø Siegpunkte (Gesamt)': '-'
            })

    df = pd.DataFrame(stats_list)
    # Sort by times played
    df = df.sort_values(by='Gespielt (Spiele)', ascending=False).reset_index(drop=True)
    return df


def df_to_excel(df_dict):
    """Exportiert mehrere DataFrames in eine Excel-Datei mit mehreren Blättern."""
    output = BytesIO()
    # Use openpyxl engine for better compatibility
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            if df is not None and not df.empty:
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
    st.subheader("Spieler eingeben")
    with st.form("player_form"):
        player_names = []
        # Create input fields for player names
        for i in range(NUM_PLAYERS):
            player_names.append(st.text_input(f"Name Spieler {i+1}", key=f"p{i}_name"))

        submitted = st.form_submit_button("Turnier starten")
        if submitted:
            # Check if exactly NUM_PLAYERS unique names were entered
            unique_names = set(filter(None, player_names)) # Filter out empty strings
            if len(unique_names) == NUM_PLAYERS:
                # Initialize player data structure
                st.session_state.players = [
                    {
                        'id': i,
                        'name': name,
                        'total_tp': 0,
                        'total_vp': 0, # NEU: Gesamt-Siegpunkte
                        'wins': 0,     # NEU: Anzahl Siege (1. Plätze)
                        'last_vp': 0,
                        'played_factions': [],
                        'played_factions_str': '',
                        'total_placement_sum': 0,
                        'games_played': 0
                     } for i, name in enumerate(unique_names) # Use unique_names to ensure correct count
                ]
                # Calculate initial turn order
                st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)
                st.session_state.initialized = True
                st.rerun() # Rerun script to show main app layout
            else:
                # Show error if names are missing or not unique
                st.error(f"Bitte {NUM_PLAYERS} eindeutige Spielernamen eingeben.")

# --- App Hauptteil (Nach Spieler-Setup) ---
if st.session_state.initialized:

    # Calculate the next turn order based on current standings
    st.session_state.next_turn_order_names = calculate_next_turn_order(st.session_state.players)

    # Define layout columns
    col1, col2 = st.columns([2, 3]) # Adjust column width ratio if needed

    with col1:
        st.subheader("📊 Aktuelle Rangliste")
        # Generate and display the standings DataFrame
        standings_df = generate_standings_df(st.session_state.players)
        # Display DataFrame WITHOUT highlighting
        st.dataframe(standings_df, use_container_width=True, hide_index=True) # Änderung hier: .style.apply() entfernt

        st.subheader("📈 Punkteentwicklung (Turnierpunkte)")
        # Display plot only if games have been played
        if not st.session_state.games:
             st.info("Noch keine Spiele gespielt, um die Entwicklung anzuzeigen.")
        else:
            # Generate data for the plot
            plot_df = generate_plot_data(st.session_state.games, st.session_state.players)
            if not plot_df.empty:
                # Create and display the line chart using Plotly Express
                fig = px.line(plot_df, x='Spiel', y='Kumulierte Turnierpunkte', color='Spieler',
                              markers=True, title="Turnierpunkte über die Zeit",
                              labels={'Spiel': 'Nach Spiel Nr.', 'Kumulierte Turnierpunkte': 'Turnierpunkte'})
                # Ensure all game numbers are shown as ticks on the x-axis
                fig.update_layout(xaxis_tickmode = 'linear')
                st.plotly_chart(fig, use_container_width=True)


    with col2:
        st.subheader(f"📜 Spiel {len(st.session_state.games) + 1} protokollieren")

        # Display warnings for repeated faction usage if any occurred in the last submission
        if st.session_state.show_faction_warning:
             for msg in st.session_state.warning_messages:
                 st.warning(msg)
             # Reset warning flag and messages after displaying
             st.session_state.show_faction_warning = False
             st.session_state.warning_messages = []


        # Form for logging a new game
        with st.form("game_log_form"):
            st.write("**Nächste Zugreihenfolge (niedrigste TP zuerst):**")
            # Display the calculated turn order for the upcoming game
            st.write(" → ".join(st.session_state.next_turn_order_names))

            # Selectbox for choosing the map
            selected_map = st.selectbox("Karte auswählen", MAPS, key=f"map_{len(st.session_state.games)}")

            game_results_input = []
            faction_warning_check = [] # List to collect warnings for this submission

            # Input fields for each player based on the calculated turn order
            for i, player_name in enumerate(st.session_state.next_turn_order_names):
                st.markdown(f"**{i+1}. {player_name}**")
                # Use columns for better layout of faction and VP input
                input_cols = st.columns(2)
                with input_cols[0]:
                    # Get player data to check previously played factions
                    player_data = get_player_data_by_name(player_name)
                    played_before = player_data['played_factions'] if player_data else []
                    # Use dynamic keys for widgets inside loops to prevent state issues
                    selected_faction = st.selectbox(f"Fraktion für {player_name}", FACTIONS, key=f"faction_{len(st.session_state.games)}_{player_name}")
                    # Check if the selected faction has been played before by this player (for warning)
                    if selected_faction in played_before:
                         faction_warning_check.append(f"Spieler **{player_name}** hat die Fraktion **{selected_faction}** bereits gespielt!")

                with input_cols[1]:
                     # Number input for victory points
                     vp = st.number_input(f"Siegpunkte (VP) für {player_name}", min_value=0, step=1, key=f"vp_{len(st.session_state.games)}_{player_name}")

                # Store the input data for this player
                game_results_input.append({'name': player_name, 'faction': selected_faction, 'vp': vp})

            # Submit button for the form
            log_game_button = st.form_submit_button("Spiel speichern")

            if log_game_button:
                # --- NEUE VALIDIERUNG HINZUGEFÜGT ---
                # 1. Prüfen, ob Fraktionen innerhalb DIESES Spiels doppelt vorkommen (BLOCKIERENDER FEHLER)
                selected_factions_this_game = [result['faction'] for result in game_results_input]
                if len(selected_factions_this_game) != len(set(selected_factions_this_game)):
                    # Finde die doppelt vorkommenden Fraktionen für eine spezifischere Fehlermeldung
                    counts = Counter(selected_factions_this_game)
                    duplicates = [faction for faction, count in counts.items() if count > 1]
                    st.error(f"Fehler: Die Fraktion(en) **{', '.join(duplicates)}** wurde(n) mehrfach ausgewählt. Jede Fraktion darf pro Spiel nur einmal vorkommen. Bitte korrigieren.")
                    # Hier anhalten, nicht speichern oder neu laden
                else:
                    # --- Nur fortfahren, wenn keine Duplikate im Spiel gefunden wurden ---

                    # 2. Nicht-blockierende Warnungen für Spieler setzen, die Fraktionen über Spiele hinweg wiederholen
                    if faction_warning_check:
                        st.session_state.show_faction_warning = True
                        st.session_state.warning_messages = faction_warning_check

                    # 3. Ergebnisse sortieren und Turnierpunkte berechnen
                    sorted_results = sorted(game_results_input, key=lambda x: x['vp'], reverse=True)

                    # Spiel-Log Eintrag vorbereiten
                    game_log_entry = {
                        'game_number': len(st.session_state.games) + 1,
                        'map': selected_map,
                        'turn_order': copy.deepcopy(st.session_state.next_turn_order_names), # Wichtig: Aktuelle Reihenfolge speichern
                        'results': []
                    }

                    # 4. Ränge, TP berechnen und Spielerstatistiken aktualisieren
                    for rank, result in enumerate(sorted_results, 1):
                        player_name = result['name']
                        # Spielerdaten im State finden
                        player_data = get_player_data_by_name(player_name)
                        if player_data:
                            # Turnierpunkte basierend auf Rang zuweisen
                            tp = TOURNAMENT_POINTS_MAP.get(rank, 0)
                            result['rank'] = rank
                            result['tp'] = tp
                            # Detailliertes Ergebnis zum Spiel-Log hinzufügen
                            game_log_entry['results'].append(result)

                            # Spielerstatistiken im State aktualisieren
                            player_data['total_tp'] += tp
                            player_data['total_vp'] += result['vp'] # Gesamt-VP addieren
                            player_data['last_vp'] = result['vp'] # Letzte Spiel VP für Tiebreaker aktualisieren
                            # Gespielte Fraktion hinzufügen (falls noch nicht vorhanden)
                            if result['faction'] not in player_data['played_factions']:
                                 player_data['played_factions'].append(result['faction'])
                            # String-Repräsentation der gespielten Fraktionen aktualisieren
                            player_data['played_factions_str'] = ", ".join(player_data['played_factions'])
                            # Summe der Platzierungen und Anzahl Spiele für Durchschnitt aktualisieren
                            player_data['total_placement_sum'] += rank
                            player_data['games_played'] += 1
                            # Siege zählen, wenn Spieler 1. wurde
                            if rank == 1:
                                player_data['wins'] += 1 # Siege zählen


                    # 5. Abgeschlossenen Spiel-Log Eintrag zur Liste hinzufügen
                    st.session_state.games.append(game_log_entry)

                    # 6. Skript neu laden, um die UI sofort zu aktualisieren (inkl. Warnungen, falls gesetzt)
                    st.rerun()
                # Ende des Else-Blocks (wird nur ausgeführt, wenn keine Duplikate im Spiel)

        # --- Zusätzliche Auswertungen ---
        if st.session_state.games: # Show stats only if games exist
            st.subheader("📊 Fraktions-Statistiken")
            faction_stats_df = calculate_faction_stats(st.session_state.games, FACTIONS)
            st.dataframe(faction_stats_df, hide_index=True, use_container_width=True)

            st.subheader("🗺️ Karten-Statistiken")
            map_stats_df = calculate_map_stats(st.session_state.games, MAPS)
            st.dataframe(map_stats_df, hide_index=True, use_container_width=True)


        # --- Spielprotokoll anzeigen ---
        st.subheader("📖 Spielprotokoll")
        # Display game logs only if games exist
        if not st.session_state.games:
            st.info("Noch keine Spiele protokolliert.")
        else:
            # Display games in reverse chronological order (most recent first)
            for game in reversed(st.session_state.games):
                # Use an expander for each game log entry
                expander_title = (f"Spiel {game['game_number']} - Karte: {game['map']} "
                                  f"(Zugreihenfolge: {' → '.join(game['turn_order'])})")
                with st.expander(expander_title):
                    # Create DataFrame for the results of this game
                    results_df = pd.DataFrame(game['results'])
                    # Select and rename columns for display
                    results_df = results_df[['rank', 'name', 'faction', 'vp', 'tp']]
                    results_df.columns = ['Platz', 'Spieler', 'Fraktion', 'Siegpunkte (VP)', 'Turnierpunkte (TP)']
                    # Display the game results table
                    st.dataframe(results_df, hide_index=True, use_container_width=True)


    # --- Export Funktion ---
    st.sidebar.title("Export")
    # Allow export only if the app is initialized and games have been played
    if st.session_state.initialized and st.session_state.games:

        # Prepare data for Excel export
        excel_data = {}
        # 1. Standings
        standings_export_df = generate_standings_df(st.session_state.players)
        # Remove styling for export
        excel_data["Rangliste"] = standings_export_df

        # 2. Faction Stats
        faction_stats_export_df = calculate_faction_stats(st.session_state.games, FACTIONS)
        excel_data["Fraktions-Statistik"] = faction_stats_export_df

        # 3. Map Stats
        map_stats_export_df = calculate_map_stats(st.session_state.games, MAPS)
        excel_data["Karten-Statistik"] = map_stats_export_df

        # 4. Detailed Game Logs
        all_games_list = []
        for game in st.session_state.games:
             turn_order_str = ", ".join(game['turn_order'])
             for result in game['results']:
                 all_games_list.append({
                     "Spiel Nr": game['game_number'],
                     "Karte": game['map'],
                     "Spieler": result['name'],
                     "Fraktion": result['faction'],
                     "Siegpunkte (VP)": result['vp'],
                     "Platz": result['rank'],
                     "Turnierpunkte (TP)": result['tp'],
                     "Zugreihenfolge (Spiel)": turn_order_str
                 })
        if all_games_list:
            # Create DataFrame from the list of game results
            excel_data["Spielprotokolle"] = pd.DataFrame(all_games_list)


        # Generate Excel file in memory
        excel_bytes = df_to_excel(excel_data)

        # Add download button to sidebar
        st.sidebar.download_button(
            label="💾 Turnierdaten als Excel exportieren",
            data=excel_bytes,
            file_name="root_turnier_ergebnisse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    elif st.session_state.initialized:
        # Show info message if no games played yet
        st.sidebar.info("Noch keine Spiele gespielt zum Exportieren.")

