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

    seed_pattern = re.compile(r"\d+(st|nd|rd|th)$", re.IGNORECASE)
    qgame_pattern = re.compile(r"Q\d+", re.IGNORECASE)

    START_ROW = 6  # QTR data begins on row 6

    for i in range(START_ROW, len(rows) - 4):
        row_q = rows[i]

        # Detect QTR row (Row C)
        if not any(qgame_pattern.match(cell.strip()) for cell in row_q):
            continue

        # Surrounding rows
        row_top_seed = rows[i - 2]        # Row A
        row_top_team = rows[i - 1]        # Row B
        row_bottom_team = rows[i + 1]     # Row D
        row_location = rows[i + 2]        # Row E

        # Scan across columns in groups of 3
        for col in range(0, len(row_q), 3):
            seed_bottom = row_q[col].strip()
            qgame = row_q[col + 1].strip() if col + 1 < len(row_q) else ""

            if not qgame_pattern.match(qgame):
                continue

            seed_top = row_top_seed[col].strip() if col < len(row_top_seed) else ""
            team_top = row_top_team[col].strip() if col < len(row_top_team) else ""
            team_bottom = row_bottom_team[col].strip() if col < len(row_bottom_team) else ""

            # date/time/location
            date = row_top_team[col + 1].strip() if col + 1 < len(row_top_team) else ""
            time = row_bottom_team[col + 1].strip() if col + 1 < len(row_bottom_team) else ""
            location = row_location[col].strip() if col < len(row_location) else ""

            # Validate seeds
            if not seed_pattern.search(seed_top):
                continue
            if not seed_pattern.search(seed_bottom):
                continue

            # Extract division name
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


def parse_semi_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    sf_pattern = re.compile(r"SF\d+", re.IGNORECASE)

    START_ROW = 6
    i = START_ROW

    while i < len(rows) - 4:
        row_seed_top = rows[i]
        row_team_top = rows[i + 1]
        row_sf = rows[i + 2]
        row_team_bottom = rows[i + 3]
        row_location = rows[i + 4]

        col = 1  # semis always in first 3 columns

        # Detect SF#
        sf_cell = row_sf[col + 1].strip()
        if not sf_pattern.match(sf_cell):
            i += 1
            continue

        sf_game = sf_cell

        # Extract seeds & teams
        higher_seed = row_seed_top[col].strip()
        higher_team = row_team_top[col].strip()

        lower_seed = row_sf[col].strip()
        lower_team = row_team_bottom[col].strip()  # may be blank

        # Extract date/time/location
        date = row_team_top[col + 1].strip()
        time = row_team_bottom[col + 1].strip()
        location = row_location[col].strip()

        # Extract division
        division = re.sub(r"\s*\d+(st|nd|rd|th)$", "", higher_seed, flags=re.IGNORECASE)
        division = division.replace(" ", "")

        results.append({
            "division": division,
            "higher_seed_label": higher_seed,
            "higher_team": higher_team,
            "lower_seed_label": lower_seed,
            "lower_team": lower_team,
            "game": sf_game,
            "date": date,
            "time": time,
            "location": location,
        })

        i += 5  # next semifinal block

    return results


def parse_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    fgame_pattern = re.compile(r"F\d+", re.IGNORECASE)

    START_ROW = 6
    i = START_ROW

    while i < len(rows) - 5:
        row_cup = rows[i]
        row_team_top = rows[i + 1]
        row_seed_top = rows[i + 2]
        row_team_bottom = rows[i + 3]
        row_seed_bottom = rows[i + 4]
        row_location = rows[i + 5]

        col = 1  # finals bracket always in first 3 columns

        # Detect F#
        f_cell = row_seed_top[col + 1].strip()
        if not fgame_pattern.match(f_cell):
            i += 1
            continue

        fgame = f_cell

        # Extract seeds & teams
        higher_team = row_team_top[col].strip()
        higher_seed = row_seed_top[col].strip()

        lower_team = row_team_bottom[col].strip()
        lower_seed = row_seed_bottom[col].strip()

        # Extract date/time/location
        date = row_team_top[col + 1].strip()
        time = row_team_bottom[col + 1].strip()
        location = row_location[col].strip()

        # Extract division from seed label
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
            "date": date,
            "time": time,
            "location": location,
        })

        i += 6  # move to next finals block

    return results


def build_brackets(qtr_data, semi_data, final_data):
    brackets = {}

    # Index semis and finals by division
    semis_by_div = {}
    for s in semi_data:
        semis_by_div.setdefault(s["division"], []).append(s)

    finals_by_div = {}
    for f in final_data:
        finals_by_div.setdefault(f["division"], []).append(f)

    # Group QTR by division
    qtrs_by_div = {}
    for q in qtr_data:
        qtrs_by_div.setdefault(q["division"], []).append(q)

    # 1. Divisions where HOLA appears in QTR
    for division, qgames in qtrs_by_div.items():
        semis = semis_by_div.get(division, [])
        finals = finals_by_div.get(division, [])

        semi = semis[0] if semis else None
        final = finals[0] if finals else None

        # Find the Q-game with HOLA
        hola_q = None
        for q in qgames:
            if "HOLA" in q["higher_team"] or "HOLA" in q["lower_team"]:
                hola_q = q
                break
        if not hola_q:
            continue

        bracket = {}

        # Quarterfinal (HOLA is in this game by definition)
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
            "date": hola_q.get("date", ""),
            "time": hola_q.get("time", ""),
            "location": hola_q.get("location", ""),
            "status": "highlight",  # HOLA reached QTR
        }

        # Semifinal
        if semi:
            # If HOLA appears in semi teams, highlight; otherwise dim
            semi_has_hola = (
                "HOLA" in semi.get("higher_team", "") or
                "HOLA" in semi.get("lower_team", "")
            )
            bracket["semifinal"] = {
                "game": semi["game"],
                "team1": {
                    "seed": semi["higher_seed_label"],
                    "name": semi["higher_team"],
                },
                "team2": {
                    "seed": semi["lower_seed_label"],
                    "name": semi["lower_team"],
                },
                "date": semi.get("date", ""),
                "time": semi.get("time", ""),
                "location": semi.get("location", ""),
                "status": "highlight" if semi_has_hola else "dim",
            }
        else:
            bracket["semifinal"] = {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "date": "",
                "time": "",
                "location": "",
                "status": "dim",
            }

        # Final
        if final:
            final_has_hola = (
                "HOLA" in final.get("higher_team", "") or
                "HOLA" in final.get("lower_team", "")
            )
            bracket["final"] = {
                "game": final["game"],
                "team1": {
                    "seed": final["higher_seed_label"],
                    "name": final["higher_team"],
                },
                "team2": {
                    "seed": final["lower_seed_label"],
                    "name": final["lower_team"],
                },
                "date": final.get("date", ""),
                "time": final.get("time", ""),
                "location": final.get("location", ""),
                "cup_name": final.get("cup_name", ""),
                "status": "highlight" if final_has_hola else "dim",
            }
        else:
            bracket["final"] = {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "date": "",
                "time": "",
                "location": "",
                "cup_name": "",
                "status": "dim",
            }

        brackets[division] = bracket

    # 2. Divisions where HOLA appears ONLY in semis
    for division, semis in semis_by_div.items():
        if division in brackets:
            continue

        hola_in_semi = any(
            "HOLA" in s.get("higher_team", "") or
            "HOLA" in s.get("lower_team", "")
            for s in semis
        )
        if not hola_in_semi:
            continue

        semi = semis[0]
        finals = finals_by_div.get(division, [])
        final = finals[0] if finals else None

        bracket = {
            "quarterfinal": {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "date": "",
                "time": "",
                "location": "",
                "status": "dim",
            },
            "semifinal": {
                "game": semi["game"],
                "team1": {
                    "seed": semi["higher_seed_label"],
                    "name": semi["higher_team"],
                },
                "team2": {
                    "seed": semi["lower_seed_label"],
                    "name": semi["lower_team"],
                },
                "date": semi.get("date", ""),
                "time": semi.get("time", ""),
                "location": semi.get("location", ""),
                "status": "highlight",
            },
        }

        if final:
            final_has_hola = (
                "HOLA" in final.get("higher_team", "") or
                "HOLA" in final.get("lower_team", "")
            )
            bracket["final"] = {
                "game": final["game"],
                "team1": {
                    "seed": final["higher_seed_label"],
                    "name": final["higher_team"],
                },
                "team2": {
                    "seed": final["lower_seed_label"],
                    "name": final["lower_team"],
                },
                "date": final.get("date", ""),
                "time": final.get("time", ""),
                "location": final.get("location", ""),
                "cup_name": final.get("cup_name", ""),
                "status": "highlight" if final_has_hola else "dim",
            }
        else:
            bracket["final"] = {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "date": "",
                "time": "",
                "location": "",
                "cup_name": "",
                "status": "dim",
            }

        brackets[division] = bracket

    # 3. Divisions where HOLA appears ONLY in finals
    for division, finals in finals_by_div.items():
        if division in brackets:
            continue

        hola_in_final = any(
            "HOLA" in f["higher_team"] or "HOLA" in f["lower_team"]
            for f in finals
        )
        if not hola_in_final:
            continue

        final = finals[0]

        bracket = {
            "quarterfinal": {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "date": "",
                "time": "",
                "location": "",
                "status": "dim",
            },
            "semifinal": {
                "game": "",
                "team1": {"seed": "", "name": ""},
                "team2": {"seed": "", "name": ""},
                "date": "",
                "time": "",
                "location": "",
                "status": "dim",
            },
            "final": {
                "game": final["game"],
                "team1": {
                    "seed": final["higher_seed_label"],
                    "name": final["higher_team"],
                },
                "team2": {
                    "seed": final["lower_seed_label"],
                    "name": final["lower_team"],
                },
                "date": final.get("date", ""),
                "time": final.get("time", ""),
                "location": final.get("location", ""),
                "cup_name": final.get("cup_name", ""),
                "status": "highlight",
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

        //
        // QUARTERFINAL
        //
        const q = document.createElement('div');
        q.className = 'round round-q';
        if (bracket.quarterfinal.game) {
          const metaClass = getMetaClass(bracket.quarterfinal.date, bracket.quarterfinal.time);
          const t1 = bracket.quarterfinal.team1.name
            ? bracket.quarterfinal.team1.seed + " " + bracket.quarterfinal.team1.name
            : bracket.quarterfinal.team1.seed;
          const t2 = bracket.quarterfinal.team2.name
            ? bracket.quarterfinal.team2.seed + " " + bracket.quarterfinal.team2.name
            : bracket.quarterfinal.team2.seed;

          q.innerHTML = `
            <div class="match match-hola ${bracket.quarterfinal.status}">
              <div class="game-label">${bracket.quarterfinal.game} • Quarterfinal</div>
              <div class="teams">
                <div>${t1}</div>
                <div>${t2}</div>
              </div>
              <div class="meta ${metaClass}">
                <span>📍 ${bracket.quarterfinal.location || "—"}</span>
                <span>🕒 ${bracket.quarterfinal.date || "—"} ${bracket.quarterfinal.time || ""}</span>
              </div>
            </div>
          `;
        }
        wrapper.appendChild(q);

        //
        // SEMIFINAL
        //
        const s = document.createElement('div');
        s.className = 'round round-s';
        if (bracket.semifinal.game) {
          const metaClass = getMetaClass(bracket.semifinal.date, bracket.semifinal.time);
          const t1 = bracket.semifinal.team1.name
            ? bracket.semifinal.team1.name
            : bracket.semifinal.team1.seed;
          const t2 = bracket.semifinal.team2.name
            ? bracket.semifinal.team2.name
            : bracket.semifinal.team2.seed;

          s.innerHTML = `
            <div class="match ${bracket.semifinal.status}">
              <div class="game-label">${bracket.semifinal.game} • Semifinal</div>
              <div class="teams">
                <div>${t1}</div>
                <div>${t2}</div>
              </div>
              <div class="meta ${metaClass}">
                <span>📍 ${bracket.semifinal.location || "—"}</span>
                <span>🕒 ${bracket.semifinal.date || "—"} ${bracket.semifinal.time || ""}</span>
              </div>
            </div>
          `;
        }
        wrapper.appendChild(s);

        //
        // FINAL
        //
        const f = document.createElement('div');
        f.className = 'round round-f';
        if (bracket.final.game) {
          const metaClass = getMetaClass(bracket.final.date, bracket.final.time);
          const t1 = bracket.final.team1.name
            ? bracket.final.team1.name
            : bracket.final.team1.seed;
          const t2 = bracket.final.team2.name
            ? bracket.final.team2.name
            : bracket.final.team2.seed;

          f.innerHTML = `
            <div class="match match-final ${bracket.final.status}">
              <div class="game-label">${bracket.final.game} • Final${bracket.final.cup_name ? " • " + bracket.final.cup_name : ""}</div>
              <div class="teams">
                <div>${t1}</div>
                <div>${t2}</div>
              </div>
              <div class="meta ${metaClass}">
                <span>📍 ${bracket.final.location || "—"}</span>
                <span>🕒 ${bracket.final.date || "—"} ${bracket.final.time || ""}</span>
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
  transition: opacity 0.2s ease, box-shadow 0.2s ease, transform 0.2s ease;
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

/* Highlight / dim states */
.match.highlight {
  opacity: 1;
  box-shadow: 0 0 0 2px #ffd54f;
  transform: translateY(-2px);
}

.match.dim {
  opacity: 0.35;
}

/* Game label */
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

/* Meta (date/time/location) */
.meta {
  margin-top: 0.5rem;
  font-size: 0.8rem;
  display: flex;
  flex-direction: column;
  gap: 0.15rem;
}

/* Date/time states */
.meta-upcoming {
  color: #a5ff9e;
  font-weight: 600;
}

.meta-today {
  color: #64b5f6;
  font-weight: 600;
}

.meta-past {
  color: #777;
  opacity: 0.7;
}

/* Connectors */
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
    print("QTR COUNT:", len(qtr_data))
    print("FIRST QTR ROW:", qtr_data[0] if qtr_data else "NONE")

    semi_data = parse_semi_finals(semis)
    print("SEMI COUNT:", len(semi_data))

    final_data = parse_finals(finals)
    print("FINAL COUNT:", len(final_data))

    brackets = build_brackets(qtr_data, semi_data, final_data)
    print("BRACKETS OBJECT:", brackets)

    write_json(brackets)
    write_html()
    write_css()


if __name__ == "__main__":
    main()
