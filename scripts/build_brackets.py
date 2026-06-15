import pandas as pd
from pathlib import Path
import json
import re
 
DATA_PATH = Path("data/2026_Playoff_Schedule_Final.xlsx")
DOCS_DIR = Path("docs")

SEED_PATTERN = re.compile(r"(.*)\s(\d+)(st|nd|rd|th)$")
QGAME_PATTERN = re.compile(r"^Q\d+$")
SGAME_PATTERN = re.compile(r"^S\d+$")
FGAME_PATTERN = re.compile(r"^F\d+$")


def load_sheets():
    xls = pd.ExcelFile(DATA_PATH)
    qtr = pd.read_excel(xls, "QTR Finals")
    semis = pd.read_excel(xls, "Semi Finals")
    finals = pd.read_excel(xls, "Finals")
    return qtr, semis, finals



def parse_qtr_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    # Skip header junk — real data starts around row 4
    start_row = 4

    i = start_row
    while i < len(rows) - 3:
        row_seed = rows[i]
        row_team = rows[i + 1]
        row_seed_low = rows[i + 2]
        row_team_low = rows[i + 3]

        # Scan across columns in groups of 3: [seed][Q#][blank]
        for col in range(1, len(row_seed), 3):
            seed_label = row_seed[col].strip()
            qgame = row_seed[col + 1].strip() if col + 1 < len(row_seed) else ""
            team = row_team[col].strip() if col < len(row_team) else ""

            seed_label_low = row_seed_low[col].strip()
            qgame_low = row_seed_low[col + 1].strip() if col + 1 < len(row_seed_low) else ""
            team_low = row_team_low[col].strip() if col < len(row_team_low) else ""

            # Must have a Q-game number like Q1, Q2, Q12, etc.
            if not QGAME_PATTERN.match(qgame):
                continue

            # Must have a valid seed label ending in st/nd/rd/th
            if not re.search(r"\d+(st|nd|rd|th)$", seed_label, re.IGNORECASE):
                continue

            # Extract division name from seed label
            division = re.sub(r"\s*\d+(st|nd|rd|th)$", "", seed_label, flags=re.IGNORECASE).strip()
            division = division.replace(" ", "")  # normalize B 6.4 → B6.4

            results.append({
                "division": division,
                "higher_seed_label": seed_label,
                "higher_team": team,
                "lower_seed_label": seed_label_low,
                "lower_team": team_low,
                "qgame": qgame,
            })

        i += 4  # move to next block

    return results

def parse_semi_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []
    current_division = None

    for r in range(len(rows)):
        row = rows[r]

        if any(k in row[0] for k in ["boys", "girls", "B6", "G6", "PG"]):
            current_division = row[0].strip()

        for c, cell in enumerate(row):
            cell = cell.strip()
            if SGAME_PATTERN.match(cell):
                sgame = cell
                field = f"Field {c + 1}"
                time = ""
                location = "Bridgewater"

                # time in same or next row
                for look in rows[r:r + 3]:
                    for item in look:
                        t = item.strip()
                        if "AM" in t or "PM" in t:
                            time = t

                winners = []
                for look in rows[r:r + 4]:
                    for item in look:
                        t = item.strip()
                        if "Winner Q" in t:
                            winners.append(t)

                if not winners:
                    winners = ["TBD", "TBD"]

                results.append({
                    "division": current_division,
                    "sgame": sgame,
                    "winners": winners,
                    "field": field,
                    "time": time,
                    "location": location,
                })

    return results


def parse_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    # column headers for fields
    header_row = rows[0]
    field_names = []
    for c, cell in enumerate(header_row):
        if "Field" in cell:
            field_names.append((c, cell.strip()))
    if not field_names:
        field_names = [(i, f"Field {i + 1}") for i in range(len(header_row))]

    for r in range(1, len(rows)):
        row = rows[r]
        for c, cell in enumerate(row):
            cell = cell.strip()
            if FGAME_PATTERN.match(cell):
                fgame = cell

                # find field by column
                field = "Field"
                for idx, name in field_names:
                    if idx == c:
                        field = name
                        break

                # look up for division/cup name
                division = ""
                for up in range(r - 1, -1, -1):
                    text = rows[up][c].strip()
                    if text and not FGAME_PATTERN.match(text) and "Seed" not in text and "AM" not in text and "PM" not in text:
                        division = text
                        break

                # time in same or previous row
                time = ""
                for look_r in [r, r - 1]:
                    if 0 <= look_r < len(rows):
                        for item in rows[look_r]:
                            t = item.strip()
                            if "AM" in t or "PM" in t:
                                time = t

                # teams around F-game cell
                team1 = ""
                team2 = ""
                # assume team1 above, team2 below in same column or nearby
                if r - 1 >= 0:
                    team1 = rows[r - 1][c].strip()
                if r + 1 < len(rows):
                    team2 = rows[r + 1][c].strip()

                results.append({
                    "division": division,
                    "fgame": fgame,
                    "team1": team1,
                    "team2": team2,
                    "field": field,
                    "time": time,
                    "location": "Bridgewater",
                })

    return results


def build_brackets(qtr_data, semi_data, final_data):
    brackets = {}

    # index semis and finals by division
    semis_by_div = {}
    for s in semi_data:
        semis_by_div.setdefault(s["division"], []).append(s)

    finals_by_div = {}
    for f in final_data:
        finals_by_div.setdefault(f["division"], []).append(f)

    # group QTR by division
    qtrs_by_div = {}
    for q in qtr_data:
        qtrs_by_div.setdefault(q["division"], []).append(q)

    for division, qgames in qtrs_by_div.items():
        # find semis that reference these Q-games
        semis = semis_by_div.get(division, [])
        finals = finals_by_div.get(division, [])

        # pick first semi and final for path (simplified)
        semi = semis[0] if semis else None
        final = finals[0] if finals else None

        # find HOLA in Q-games
        hola_q = None
        for q in qgames:
            if "HOLA" in q["higher_team"] or "HOLA" in q["lower_team"]:
                hola_q = q
                break
        if not hola_q:
            continue  # skip divisions without HOLA in QTR

        bracket = {}

        bracket["quarterfinal"] = {
            "game": hola_q["qgame"],
            "team1": {
                "seed": hola_q["higher_seed_label"],
                "name": hola_q["higher_team"],
            },
            "team2": {
                "seed": hola_q["lower_seed_label"],
                "name": hola_q["lower_team"],
            },
            "location": "Bridgewater F?",
            "time": "",
        }

        if semi:
            bracket["semifinal"] = {
                "game": semi["sgame"],
                "from": semi["winners"],
                "location": f"{semi['location']} {semi['field']}",
                "time": semi["time"],
            }
        else:
            bracket["semifinal"] = {
                "game": "",
                "from": ["TBD", "TBD"],
                "location": "",
                "time": "",
            }

        if final:
            bracket["final"] = {
                "game": final["fgame"],
                "from": [final["team1"] or "Winner Semi", final["team2"] or "Winner other Semi"],
                "location": f"{final['location']} {final['field']}",
                "time": final["time"],
            }
        else:
            bracket["final"] = {
                "game": "",
                "from": ["TBD", "TBD"],
                "location": "",
                "time": "",
            }

        brackets[division] = bracket

    # also handle divisions where HOLA appears only in semis or finals
    for division, semis in semis_by_div.items():
        if division in brackets:
            continue
        hola_in_semi = any("HOLA" in " ".join(s["winners"]) for s in semis)
        if not hola_in_semi:
            continue
        finals = finals_by_div.get(division, [])
        semi = semis[0]
        final = finals[0] if finals else None

        bracket = {
            "quarterfinal": {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "location": "",
                "time": "",
            },
            "semifinal": {
                "game": semi["sgame"],
                "from": semi["winners"],
                "location": f"{semi['location']} {semi['field']}",
                "time": semi["time"],
            },
        }

        if final:
            bracket["final"] = {
                "game": final["fgame"],
                "from": [final["team1"] or "Winner Semi", final["team2"] or "Winner other Semi"],
                "location": f"{final['location']} {final['field']}",
                "time": final["time"],
            }
        else:
            bracket["final"] = {
                "game": "",
                "from": ["TBD", "TBD"],
                "location": "",
                "time": "",
            }

        brackets[division] = bracket

    for division, finals in finals_by_div.items():
        if division in brackets:
            continue
        hola_in_final = any("HOLA" in (f["team1"] + f["team2"]) for f in finals)
        if not hola_in_final:
            continue
        final = finals[0]

        bracket = {
            "quarterfinal": {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "location": "",
                "time": "",
            },
            "semifinal": {
                "game": "",
                "from": ["TBD", "TBD"],
                "location": "",
                "time": "",
            },
            "final": {
                "game": final["fgame"],
                "from": [final["team1"], final["team2"]],
                "location": f"{final['location']} {final['field']}",
                "time": final["time"],
            },
        }

        brackets[division] = bracket

    return brackets


def write_json(brackets):
    DOCS_DIR.mkdir(exist_ok=True)
    with open(DOCS_DIR / "brackets.json", "w", encoding="utf-8") as f:
        json.dump(brackets, f, indent=2)


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
    <p>Visual paths to the championship for all HOLA teams.</p>
  </header>

  <main id="brackets"></main>

  <script>
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

        const q = document.createElement('div');
        q.className = 'round round-q';
        if (bracket.quarterfinal.game) {
          q.innerHTML = `
            <div class="match match-hola">
              <div class="game-label">${bracket.quarterfinal.game} • Quarterfinal</div>
              <div class="teams">
                <div>${bracket.quarterfinal.team1.seed} ${bracket.quarterfinal.team1.name}</div>
                <div>${bracket.quarterfinal.team2.seed} ${bracket.quarterfinal.team2.name}</div>
              </div>
              <div class="meta">
                <span>📍 ${bracket.quarterfinal.location}</span>
                <span>🕒 ${bracket.quarterfinal.time}</span>
              </div>
            </div>
          `;
        }
        wrapper.appendChild(q);

        const s = document.createElement('div');
        s.className = 'round round-s';
        if (bracket.semifinal.game) {
          s.innerHTML = `
            <div class="match">
              <div class="game-label">${bracket.semifinal.game} • Semifinal</div>
              <div class="teams">
                <div>${bracket.semifinal.from[0]}</div>
                <div>${bracket.semifinal.from[1]}</div>
              </div>
              <div class="meta">
                <span>📍 ${bracket.semifinal.location}</span>
                <span>🕒 ${bracket.semifinal.time}</span>
              </div>
            </div>
          `;
        }
        wrapper.appendChild(s);

        const f = document.createElement('div');
        f.className = 'round round-f';
        if (bracket.final.game) {
          f.innerHTML = `
            <div class="match match-final">
              <div class="game-label">${bracket.final.game} • Final</div>
              <div class="teams">
                <div>${bracket.final.from[0]}</div>
                <div>${bracket.final.from[1]}</div>
              </div>
              <div class="meta">
                <span>📍 ${bracket.final.location}</span>
                <span>🕒 ${bracket.final.time}</span>
              </div>
            </div>
          `;
        }
        wrapper.appendChild(f);

        section.appendChild(wrapper);
        container.appendChild(section);
      }
    }

    loadBrackets();
  </script>
</body>
</html>
"""
    DOCS_DIR.mkdir(exist_ok=True)
    with open(DOCS_DIR / "index.html", "w", encoding="utf-8") as f:
        f.write(html)


def write_css():
    css = """body {
  font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  background: #0b1020;
  color: #f5f5f5;
  margin: 0;
  padding: 0;
}

header {
  padding: 1.5rem;
  text-align: center;
  background: #11162a;
  border-bottom: 1px solid #222a3f;
}

h1 {
  margin: 0 0 0.5rem;
}

main {
  padding: 1.5rem;
}

.division {
  margin-bottom: 3rem;
}

.division h2 {
  margin-bottom: 1rem;
  border-left: 4px solid #ffd54f;
  padding-left: 0.5rem;
}

.bracket-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 1.5rem;
  align-items: center;
}

.round {
  position: relative;
}

.match {
  background: #1a2138;
  border-radius: 8px;
  padding: 0.75rem 1rem;
  box-shadow: 0 0 0 1px #252c45;
}

.match-hola {
  border-color: #ffd54f;
  box-shadow: 0 0 0 2px #ffd54f;
  background: #2b2b10;
}

.match-final {
  border-color: #ffb300;
  box-shadow: 0 0 0 2px #ffb300;
}

.game-label {
  font-size: 0.8rem;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: #9fa8da;
  margin-bottom: 0.5rem;
}

.teams > div {
  margin-bottom: 0.25rem;
}

.meta {
  margin-top: 0.5rem;
  font-size: 0.8rem;
  color: #c5cae9;
  display: flex;
  flex-direction: column;
  gap: 0.15rem;
}

.round-q::after,
.round-s::after {
  content: '';
  position: absolute;
  right: -0.75rem;
  top: 50%;
  width: 0.75rem;
  height: 2px;
  background: #c5cae9;
}

.round-s::before {
  content: '';
  position: absolute;
  left: -0.75rem;
  top: 50%;
  width: 0.75rem;
  height: 2px;
  background: #c5cae9;
}

.round-f::before {
  content: '';
  position: absolute;
  left: -0.75rem;
  top: 50%;
  width: 0.75rem;
  height: 2px;
  background: #ffb300;
}

@media (max-width: 768px) {
  .bracket-grid {
    grid-template-columns: 1fr;
  }

  .round-q::after,
  .round-s::after,
  .round-s::before,
  .round-f::before {
    display: none;
  }
}
"""
    DOCS_DIR.mkdir(exist_ok=True)
    with open(DOCS_DIR / "style.css", "w", encoding="utf-8") as f:
        f.write(css)


def main():
    qtr, semis, finals = load_sheets()
    qtr_data = parse_qtr_finals(qtr)
    semi_data = parse_semi_finals(semis)
    final_data = parse_finals(finals)

    brackets = build_brackets(qtr_data, semi_data, final_data)
    write_json(brackets)
    write_html()
    write_css()

qtr_data = parse_qtr_finals(qtr_df)
print("QTR COUNT:", len(qtr_data))
print("FIRST QTR ROW:", qtr_data[0] if qtr_data else "NONE")


if __name__ == "__main__":
    main()
