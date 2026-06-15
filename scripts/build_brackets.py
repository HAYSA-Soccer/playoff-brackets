import pandas as pd
from pathlib import Path
import json
import re

DATA_PATH = Path("data/2026_Playoff_Schedule_Final.xlsx")
DOCS_DIR = Path("docs")

# HOLA divisions (canonical)
HOLA_DIVISIONS = {"B6.4", "G6.2", "B6.2", "B8.2", "G8.2", "B10.2", "BPG.1"}

QGAME_PATTERN = re.compile(r"Q(\d+)", re.IGNORECASE)
SGAME_PATTERN = re.compile(r"SF(\d+)", re.IGNORECASE)
FGAME_PATTERN = re.compile(r"F(\d+)", re.IGNORECASE)


# -----------------------------
#  DIVISION NORMALIZATION
# -----------------------------
def normalize_division(label):
    if not label:
        return ""

    s = label.replace(" ", "").lower()

    # Boys
    m = re.match(r"b(\d+\.\d+)", s)
    if m:
        return f"B{m.group(1)}"

    # Girls
    m = re.match(r"g(\d+\.\d+)", s)
    if m:
        return f"G{m.group(1)}"

    # Formats like "8.2boys" or "8.2girls"
    m = re.match(r"(\d+\.\d+)(boys|girls)", s)
    if m:
        prefix = "B" if m.group(2) == "boys" else "G"
        return f"{prefix}{m.group(1)}"

    # PG divisions
    if "pg" in s:
        m = re.match(r"(b|g)pg\.?(\d*)", s)
        if m:
            return f"{m.group(1).upper()}PG.{m.group(2) or '1'}"

    return label.upper()


# -----------------------------
#  LOAD SHEETS
# -----------------------------
def load_sheets():
    xls = pd.ExcelFile(DATA_PATH)
    return (
        pd.read_excel(xls, "QTR Finals"),
        pd.read_excel(xls, "Semi Finals"),
        pd.read_excel(xls, "Finals"),
    )


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

            raw_div = re.sub(r"\s*\d+(st|nd|rd|th)$", "", seed_top, flags=re.IGNORECASE)
            division = normalize_division(raw_div)

            results.append({
                "division": division,
                "higher_seed_label": seed_top,
                "higher_team": team_top,
                "lower_seed_label": seed_bottom,
                "lower_team": team_bottom,
                "qgame": qgame,
                "qnum": QGAME_PATTERN.match(qgame).group(1),
                "date": date,
                "time": time,
                "location": location,
            })

    return results


# -----------------------------
#  SEMIFINAL PARSER
# -----------------------------
def parse_semi_finals(df):
    rows = df.fillna("").astype(str).values.tolist()
    results = []

    START_ROW = 6
    i = START_ROW

    while i < len(rows) - 4:
        row_seed_top = rows[i]
        row_team_top = rows[i + 1]
        row_seed_bottom = rows[i + 2]
        row_team_bottom = rows[i + 3]
        row_location = rows[i + 4]

        col = 1

        sf_cell = row_seed_bottom[col + 1].strip()
        m = SGAME_PATTERN.match(sf_cell)
        if not m:
            i += 1
            continue

        sf_num = m.group(1)

        higher_seed = row_seed_top[col].strip()
        higher_team = row_team_top[col].strip()

        lower_seed = row_seed_bottom[col].strip()
        lower_team = row_team_bottom[col].strip()

        date = row_team_top[col + 1].strip()
        time = row_team_bottom[col + 1].strip()
        location = row_location[col].strip()

        raw_div = re.sub(r"\s*\d+(st|nd|rd|th)$", "", higher_seed, flags=re.IGNORECASE)
        division = normalize_division(raw_div)

        results.append({
            "division": division,
            "higher_seed_label": higher_seed,
            "higher_team": higher_team,
            "lower_seed_label": lower_seed,
            "lower_team": lower_team,
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

        f_num = m.group(1)

        higher_team = row_team_top[col].strip()
        higher_seed = row_seed_top[col].strip()

        lower_team = row_team_bottom[col].strip()
        lower_seed = row_seed_bottom[col].strip()

        date = row_team_top[col + 1].strip()
        time = row_team_bottom[col + 1].strip()
        location = row_location[col].strip()

        raw_div = re.sub(r"\s*\d+(st|nd|rd|th)$", "", higher_seed, flags=re.IGNORECASE)
        division = normalize_division(raw_div)

        results.append({
            "division": division,
            "cup_name": row_cup[col].strip(),
            "higher_seed_label": higher_seed,
            "higher_team": higher_team,
            "lower_seed_label": lower_seed,
            "lower_team": lower_team,
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
        bracket = {}

        qgames = qtrs_by_div.get(division, [])
        semis = semis_by_div.get(division, [])
        finals = finals_by_div.get(division, [])

        # -------------------------
        # QUARTERFINAL (only if exists)
        # -------------------------
        if qgames:
            q = qgames[0]
            bracket["quarterfinal"] = {
                "exists": True,
                "team1": q["higher_team"],
                "team2": q["lower_team"],
                "seed1": q["higher_seed_label"],
                "seed2": q["lower_seed_label"],
                "game": f"Quarterfinal Game #{q['qnum']}",
                "date": q["date"],
                "time": q["time"],
                "location": q["location"],
                "highlight": ("HOLA" in q["higher_team"] or "HOLA" in q["lower_team"]),
            }
        else:
            bracket["quarterfinal"] = {"exists": False}

        # -------------------------
        # SEMIFINAL (only if exists)
        # -------------------------
        if semis:
            s = semis[0]

            # Determine semifinal team names
            if s["higher_team"]:
                t1 = s["higher_team"]
            else:
                if qgames:
                    t1 = f"Winner of Quarterfinal Game #{qgames[0]['qnum']}"
                else:
                    t1 = ""

            if s["lower_team"]:
                t2 = s["lower_team"]
            else:
                if qgames:
                    t2 = f"Winner of Quarterfinal Game #{qgames[0]['qnum']}"
                else:
                    t2 = ""

            bracket["semifinal"] = {
                "exists": True,
                "team1": t1,
                "team2": t2,
                "seed1": s["higher_seed_label"],
                "seed2": s["lower_seed_label"],
                "game": f"Semifinal Game #{s['sf_num']}",
                "date": s["date"],
                "time": s["time"],
                "location": s["location"],
                "highlight": ("HOLA" in t1 or "HOLA" in t2),
            }
        else:
            bracket["semifinal"] = {"exists": False}

        # -------------------------
        # FINAL (only if exists)
        # -------------------------
        if finals:
            f = finals[0]

            # Determine final team names
            if f["higher_team"]:
                t1 = f["higher_team"]
            else:
                if semis:
                    t1 = f"Winner of Semifinal Game #{semis[0]['sf_num']}"
                else:
                    t1 = ""

            if f["lower_team"]:
                t2 = f["lower_team"]
            else:
                if semis:
                    t2 = f"Winner of Semifinal Game #{semis[0]['sf_num']}"
                else:
                    t2 = ""

            bracket["final"] = {
                "exists": True,
                "team1": t1,
                "team2": t2,
                "seed1": f["higher_seed_label"],
                "seed2": f["lower_seed_label"],
                "game": f"Final Game #{f['f_num']}",
                "cup_name": f["cup_name"],
                "date": f["date"],
                "time": f["time"],
                "location": f["location"],
                "highlight": ("HOLA" in t1 or "HOLA" in t2),
            }
        else:
            bracket["final"] = {"exists": False}

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
  <title>HOLA Playoff Brackets 2026</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <header>
    <h1>HOLA Playoff Brackets 2026</h1>
  </header>

  <main id="brackets"></main>

  <script>
    function getMetaClass(round) {
      if (!round.highlight) return "meta-dim";

      if (!round.date || !round.time) return "meta-dim";

      const dt = new Date(round.date + " " + round.time);
      if (isNaN(dt.getTime())) return "meta-dim";

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
          if (!round.exists) return null;

          const r = document.createElement('div');
          r.className = 'round';

          const metaClass = getMetaClass(round);

          const t1 = round.team1 || "";
          const t2 = round.team2 || "";

          r.innerHTML = `
            <div class="match ${round.highlight ? "highlight" : "dim"}">
              <div class="game-label">${round.game}</div>
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

        const q = renderRound(bracket.quarterfinal, "Quarterfinal");
        const s = renderRound(bracket.semifinal, "Semifinal");
        const f = renderRound(bracket.final, "Final");

        if (q) wrapper.appendChild(q);
        if (s) wrapper.appendChild(s);
        if (f) wrapper.appendChild(f);

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

.meta-dim {
  color: #555;
  opacity: 0.5;
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


