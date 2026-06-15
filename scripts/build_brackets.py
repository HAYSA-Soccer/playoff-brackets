import pandas as pd
from pathlib import Path
import json

DATA_PATH = Path("data/2026_Playoff_Schedule_Final.xlsx")
DOCS_DIR = Path("docs")

def load_sheets():
    xls = pd.ExcelFile(DATA_PATH)
    qtr = pd.read_excel(xls, "QTR Finals")
    semis = pd.read_excel(xls, "Semi Finals")   # ← no hyphen
    finals = pd.read_excel(xls, "Finals")
    return qtr, semis, finals

def extract_hola_brackets(qtr, semis, finals):
    # TODO: adapt to your exact column names
    brackets = {}

    # Example: 8.2 Boys HOLA path
    # You’ll map division names, game numbers, seeds, teams, locations, times.
    # This is just a placeholder structure:
    brackets["8.2 Boys"] = {
        "quarterfinal": {
            "game": "Q3",
            "team1": {"seed": 3, "name": "WHTH (Gould)"},
            "team2": {"seed": 6, "name": "HOLA (Luyo)"},
            "location": "Bridgewater F4",
            "time": "Jun 14 • 6:00 PM",
        },
        "semifinal": {
            "game": "S?",
            "from": ["Winner Q3", "Winner Q4"],
            "location": "Bridgewater F2",
            "time": "Jun 20 • 8:00 AM",
        },
        "final": {
            "game": "F15",
            "from": ["Winner S?", "Winner other semi"],
            "location": "Bridgewater F3",
            "time": "Jun 21 • 9:00 AM",
        },
    }

    # Repeat for PG Boys, 6.2 Girls, B6.4 Boys, G8.2 Girls using real data.
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

        // Build visual bracket layout
        const wrapper = document.createElement('div');
        wrapper.className = 'bracket-grid';

        // Quarterfinal
        const q = document.createElement('div');
        q.className = 'round round-q';
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
        wrapper.appendChild(q);

        // Semifinal
        const s = document.createElement('div');
        s.className = 'round round-s';
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
        wrapper.appendChild(s);

        // Final
        const f = document.createElement('div');
        f.className = 'round round-f';
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

/* Simple connector lines (visual path) */
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

/* Mobile */
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
    with open(DOCS_DIR / "style.css", "w", encoding="utf-8") as f:
        f.write(css)

def main():
    qtr, semis, finals = load_sheets()
    brackets = extract_hola_brackets(qtr, semis, finals)
    write_json(brackets)
    write_html()
    write_css()

if __name__ == "__main__":
    main()
