import pandas as pd
from pathlib import Path
import json
import re

DATA_PATH = Path("data/2026_Playoff_Schedule_Final.xlsx")
DOCS_DIR = Path("docs")

# Your HOLA divisions
HOLA_DIVISIONS = {"B6.4", "G6.2", "B6.2", "B8.2", "G8.2", "B10.2", "BPG.1"}

QGAME_PATTERN = re.compile(r"Q(\d+)", re.IGNORECASE)
SGAME_PATTERN = re.compile(r"SF(\d+)", re.IGNORECASE)
FGAME_PATTERN = re.compile(r"F(\d+)", re.IGNORECASE)


def load_sheets():
    xls = pd.ExcelFile(DATA_PATH)
    qtr = pd.read_excel(xls, "QTR Finals")
    semis = pd.read_excel(xls, "Semi Finals")
    finals = pd.read_excel(xls, "Finals")
    return qtr, semis, finals


# -----------------------------
#  QUARTERFINAL PARSER
# -----------------------------
def parse_qtr_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    seed_pattern = re.compile(r"\d+(st|nd|rd|th)$", re.IGNORECASE)

    START_ROW = 6
    for i in range(START_ROW, len(rows) - 4):
        row_q = rows[i]

        # Detect Q-game in row
        if not any(QGAME_PATTERN.match(cell.strip()) for cell in row_q):
            continue

        row_seed_top = rows[i - 2]
        row_team_top = rows[i - 1]
        row_team_bottom = rows[i + 1]
        row_location = rows[i + 2]

        for col in range(0, len(row_q), 3):
            seed_bottom = row_q[col].strip()
            qgame = row_q[col + 1].strip() if col + 1 < len(row_q) else ""

            if not QGAME_PATTERN.match(qgame):
                continue

            seed_top = row_seed_top[col].strip()
            team_top = row_team_top[col].strip()
            team_bottom = row_team_bottom[col].strip()

            date = row_team_top[col + 1].strip() if col + 1 < len(row_team_top) else ""
            time = row_team_bottom[col + 1].strip() if col + 1 < len(row_team_bottom) else ""
            location = row_location[col].strip()

            if not seed_pattern.search(seed_top):
                continue
            if not seed_pattern.search(seed_bottom):
                continue

            division = re.sub(r"\s*\d+(st|nd|rd|th)$", "", seed_top, flags=re.IGNORECASE)
            division = division.replace(" ", "")

            results.append({
                "division": division,
                "higher_seed_label": seed_top,
                "higher_team": team_top,
                "lower_seed_label": seed_bottom,
                "lower_team": team_bottom,
                "qgame": qgame,
                "date": date,
                "time": time,
                "location": location,
            })

    return results


# -----------------------------
#  SEMIFINAL PARSER (5-row blocks)
# -----------------------------
def parse_semi_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    START_ROW = 6
    i = START_ROW

    while i < len(rows) - 4:
        row_seed_top = rows[i]        # Row 6
        row_team_top = rows[i + 1]    # Row 7
        row_seed_bottom = rows[i + 2] # Row 8
        row_team_bottom = rows[i + 3] # Row 9
        row_location = rows[i + 4]    # Row 10

        col = 1  # Bracket always starts in column B

        sf_cell = row_seed_bottom[col + 1].strip()
        m = SGAME_PATTERN.match(sf_cell)
        if not m:
            i += 1
            continue

        sf_game = sf_cell
        sf_num = m.group(1)

        higher_seed = row_seed_top[col].strip()
        higher_team = row_team_top[col].strip()

        lower_seed = row_seed_bottom[col].strip()
        lower_team = row_team_bottom[col].strip()

        date = row_team_top[col + 1].strip()
        time = row_team_bottom[col + 1].strip()
        location = row_location[col].strip()

        division = re.sub(r"\s*\d+(st|nd|rd|th)$", "", higher_seed, flags=re.IGNORECASE)
        division = division.replace(" ", "")

        results.append({
            "division": division,
            "higher_seed_label": higher_seed,
            "higher_team": higher_team,
            "lower_seed_label": lower_seed,
            "lower_team": lower_team,
            "game": sf_game,
            "sf_num": sf_num,
            "date": date,
            "time": time,
            "location": location,
        })

        i += 5

    return results


# -----------------------------
#  FINALS PARSER
# -----------------------------
def parse_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    START_ROW = 6
    i = START_ROW

    while i < len(rows) - 5:
        row_cup = rows[i]
        row_team_top = rows[i + 1]
        row_seed_top = rows[i + 2]
        row_team_bottom = rows[i + 3]
        row_seed_bottom = rows[i + 4]
        row_location = rows[i + 5]

        col = 1

        f_cell = row_seed_top[col + 1].strip()
        m = FGAME_PATTERN.match(f_cell)
        if not m:
            i += 1
            continue

        fgame = f_cell
        f_num = m.group(1)

        higher_team = row_team_top[col].strip()
        higher_seed = row_seed_top[col].strip()

        lower_team = row_team_bottom[col].strip()
        lower_seed = row_seed_bottom[col].strip()

        date = row_team_top[col + 1].strip()
        time = row_team_bottom[col + 1].strip()
        location = row_location[col].strip()

        division = re.sub(r"\s*\d+(st|nd|rd|th)$", "", higher_seed, flags=re.IGNORECASE)
        division = division.replace(" ", "")

        results.append({
            "division": division,
            "cup_name": row_cup[col].strip(),
            "higher_seed_label": higher_seed,
            "higher_team": higher_team,
            "lower_seed_label": lower_seed,
            "lower_team": lower_team,
            "game": fgame,
            "f_num": f_num,
            "date": date,
            "time": time,
            "location": location,
        })

        i += 6

    return results


# -----------------------------
#  BUILD BRACKETS
# -----------------------------
def build_brackets(qtr_data, semi_data, final_data):
    brackets = {}

    qtrs_by_div = {}
    for q in qtr_data:
        qtrs_by_div.setdefault(q["division"], []).append(q)

    semis_by_div = {}
    for s in semi_data:
        semis_by_div.setdefault(s["division"], []).append(s)

    finals_by_div = {}
    for f in final_data:
        finals_by_div.setdefault(f["division"], []).append(f)

    for division in HOLA_DIVISIONS:
        qgames = qtrs_by_div.get(division, [])
        semis = semis_by_div.get(division, [])
        finals = finals_by_div.get(division, [])

        bracket = {}

        # -------------------------
        # QUARTERFINAL
        # -------------------------
        if qgames:
            q = qgames[0]
            q_num = QGAME_PATTERN.match(q["qgame"]).group(1)

            bracket["quarterfinal"] = {
                "game": q["qgame"],
                "team1": {"seed": q["higher_seed_label"], "name": q["higher_team"]},
                "team2": {"seed": q["lower_seed_label"], "name": q["lower_team"]},
                "date": q["date"],
                "time": q["time"],
                "location": q["location"],
                "status": "highlight" if "HOLA" in q["higher_team"] or "HOLA" in q["lower_team"] else "dim",
            }
        else:
            bracket["quarterfinal"] = {
                "game": "",
                "team1": {"seed": "", "name": f"Winner of Quarterfinal Game 11"},
                "team2": {"seed": "", "name": f"Winner of Quarterfinal Game 12"},
                "date": "",
                "time": "",
                "location": "",
                "status": "dim",
            }

        # -------------------------
        # SEMIFINAL
        # -------------------------
        if semis:
            s = semis[0]
            sf_num = s["sf_num"]

            t1 = s["higher_team"] or f"Winner of Quarterfinal Game {sf_num}"
            t2 = s["lower_team"] or f"Winner of Quarterfinal Game {sf_num}"

            bracket["semifinal"] = {
                "game": s["game"],
                "team1": {"seed": s["higher_seed_label"], "name": t1},
                "team2": {"seed": s["lower_seed_label"], "name": t2},
                "date": s["date"],
                "time": s["time"],
                "location": s["location"],
                "status": "highlight" if "HOLA" in t1 or "HOLA" in t2 else "dim",
            }
        else:
            bracket["semifinal"] = {
                "game": "",
                "team1": {"seed": "", "name": "Winner of Quarterfinal Game 11"},
                "team2": {"seed": "", "name": "Winner of Quarterfinal Game 12"},
                "date": "",
                "time": "",
                "location": "",
                "status": "dim",
            }

        # -------------------------
        # FINAL
        # -------------------------
        if finals:
            f = finals[0]
            f_num = f["f_num"]

            t1 = f["higher_team"] or f"Winner of Semifinal Game {f_num}"
            t2 = f["lower_team"] or f"Winner of Semifinal Game {f_num}"

            bracket["final"] = {
                "game": f["game"],
                "team1": {"seed": f["higher_seed_label"], "name": t1},
                "team2": {"seed": f["lower_seed_label"], "name": t2},
                "date": f["date"],
                "time": f["time"],
                "location": f["location"],
                "cup_name": f["cup_name"],
                "status": "highlight" if "HOLA" in t1 or "HOLA" in t2 else "dim",
            }
        else:
            bracket["final"] = {
                "game": "",
                "team1": {"seed": "", "name": "Winner of Semifinal Game 11"},
                "team2": {"seed": "", "name": "Winner of Semifinal Game 12"},
                "date": "",
                "time": "",
                "location": "",
                "cup_name": "",
                "status": "dim",
            }

        brackets[division] = bracket

    return brackets


# -----------------------------
#  WRITE JSON
# -----------------------------
def write_json(brackets):
    DOCS_DIR.mkdir(exist_ok=True)
    with open(DOCS_DIR / "brackets.json", "w", encoding="utf-8") as f:
        json.dump(brackets, f, indent=2)


# -----------------------------
#  WRITE HTML
# -----------------------------
def write_html():
    html = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>HOLA Playoff Brackets</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <header>
    <h1>HOLA Playoff Brackets 2026</h1>
  </header>

  <main id="brackets"></main>

  <script>
    function getMetaClass(dateStr, timeStr) {
      if (!dateStr || !timeStr) return "meta-past";

      const dt = new Date(dateStr + " " + timeStr);
      if (isNaN(dt.getTime())) return "meta-past";

      const now = new Date();

      const sameDay =
        dt.getFullYear() === now.getFullYear() &&
        dt.getMonth() === now.getMonth() &&
        dt.getDate() === now.getDate();

      if (dt > now) return "meta-upcoming";
      if (sameDay) return "meta-today";
      return "meta-past";
    }

    async function loadBrackets() {
      const res = await fetch('brackets.json');
      const data = await res.json();
      const container = document.getElementById('brackets');

      for (const [division, bracket] of Object.entries(data)) {
        const section = document.createElement('section');
        section.className = 'division';

        const title = document.createElement('h2');
        title.textContent = division;
        section.appendChild(title);

        const wrapper = document.createElement('div');
        wrapper.className = 'bracket-grid';

        function renderRound(round, label) {
          const r = document.createElement('div');
          r.className = 'round';

          const metaClass = getMetaClass(round.date, round.time);

          const t1 = round.team1.name || round.team1.seed;
          const t2 = round.team2.name || round.team2.seed;

          r.innerHTML = `
            <div class="match ${round.status}">
              <div class="game-label">${round.game || label}</div>
              <div class="teams">
                <div>${t1}</div>
                <div>${t2}</div>
              </div>
              <div class="meta ${metaClass}">
                <span>📍 ${round.location || "—"}</span>
                <span>🕒 ${round.date || "—"} ${round.time || ""}</span>
              </div>
            </div>
          `;
          return r;
        }

        wrapper.appendChild(renderRound(bracket.quarterfinal, "Quarterfinal"));
        wrapper.appendChild(renderRound(bracket.semifinal, "Semifinal"));
        wrapper.appendChild(renderRound(bracket.final, "Final"));

        section.appendChild(wrapper);
        container.appendChild(section);
      }
    }

    loadBrackets();
  </script>
</body>
</html>
"""
    with open(DOCS_DIR / "index.html", "w", encoding="utf-8") as f:
        f.write(html)


# -----------------------------
#  WRITE CSS
# -----------------------------
def write_css():
    css = """
body {
  font-family: system-ui, sans-serif;
  background: #0b1020;
  color: #f5f5f5;
  margin: 0;
  padding: 0;
}

header {
  padding: 1.5rem;
  text-align: center;
  background: #11162a;
}

main {
  padding: 1.5rem;
}

.division {
  margin-bottom: 3rem;
}

.bracket-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 1.5rem;
}

.match {
  background: #1a2138;
  border-radius: 8px;
  padding: 1rem;
  box-shadow: 0 0 0 1px #252c45;
}

.match.highlight {
  opacity: 1;
  box-shadow: 0 0 0 2px #ffd54f;
}

.match.dim {
  opacity: 0.35;
}

.game-label {
  font-size: 0.8rem;
  color: #9fa8da;
  margin-bottom: 0.5rem;
}

.meta-upcoming {
  color: #a5ff9e;
}

.meta-today {
  color: #64b5f6;
}

.meta-past {
  color: #777;
  opacity: 0.7;
}
"""
    with open(DOCS_DIR / "style.css", "w", encoding="utf-8") as f:
        f.write(css)


# -----------------------------
#  MAIN
# -----------------------------
def main():
    qtr, semis, finals = load_sheets()

    qtr_data = parse_qtr_finals(qtr)
    semi_data = parse_semi_finals(semis)
    final_data = parse_finals(finals)

    brackets = build_brackets(qtr_data, semi_data, final_data)

    write_json(brackets)
    write_html()
    write_css()


if __name__ == "__main__":
    main()
