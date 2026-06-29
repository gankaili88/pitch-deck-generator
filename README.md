# Consulting Pitch Deck Generator

An AI-powered tool that generates structured consulting pitch decks as downloadable PowerPoint files. Built for aspiring and current consultants who want a credible first-draft accelerator grounded in real consulting deck conventions.

> 🔗 **Live demo:** [pitch-deck-generator-gankaili.streamlit.app](https://pitch-deck-generator-gankaili.streamlit.app/)
> 📊 **Companion project:** [Excel Formula Generator (Project 1)](https://github.com/gankaili88/excel-formula-generator)
---

---

## What it does

Generates a structured 5-7 slide consulting pitch deck from a client brief. Each deck includes:

- A title slide with the engagement framing
- A structured analytical body (chosen framework)
- A strategic options slide as a 3-column comparison diagram
- A recommendation slide with rationale
- Every slide has a "so-what" title, body points, and a key takeaway bar
- Output is a fully-formatted `.pptx` file the user can open in PowerPoint or Google Slides

## Key features

- **6 analytical frameworks** — SCR, Porter's Five Forces, SWOT, McKinsey 7S, Three Horizons, MECE issue tree. The chosen framework drives both the deck structure and the AI's analytical approach.
- **Research context grounding** — users paste real client research (press releases, financial data, competitor names). The AI prioritises these specific facts over generic content.
- **Industry-aware visual themes** — 8 hand-tuned colour palettes that auto-suggest based on client name/industry (Banking Blue for banks, Energy Bold for oil & gas, Healthcare Calm for hospitals, etc.). Manual override always available.
- **3-column options diagram** — the strategic options slide renders as a visual comparison rather than bullet points, mirroring how senior consultants present strategic choices.
- **Polished PowerPoint output** — header strips, page numbers, footers, refined typography. Generated deck looks professionally produced, not like an AI artefact.

---

## Why I built it

I'm an ICAEW student preparing to join a Big Four firm as a consulting associate. After building [Project 1 — Excel Formula Generator](https://github.com/gankaili88/excel-formula-generator), I wanted to tackle a harder problem: encoding *consulting structural thinking* into prompts and producing a tangible artefact (a real `.pptx`) rather than just text output.

The blank-page problem is real in consulting. Every junior consultant has experienced staring at an empty slide before drafting. A tool that produces a structured first draft — even an imperfect one — removes that blocker.

---

## Tech stack

- **Python + Streamlit** — UI layer
- **Google Gemini 2.5 Flash API** — generation engine (free tier)
- **python-pptx** — programmatic PowerPoint generation
- **python-dotenv** — secure config management

---

## Architecture notes worth flagging

**Provider-agnostic structure.** The LLM call is isolated in one function (`call_gemini_for_deck`). Swapping Gemini for Claude, GPT, or a local model is a one-function change. Vendor lock-in is a real enterprise concern in AI deployments, and this structure reflects that.

**Structured JSON output.** The prompt instructs Gemini to return JSON with a specific schema, which is parsed in Python and rendered into PowerPoint. This is the right pattern for any structured AI output — much more reliable than parsing free text.

**Defensive parsing.** The code strips markdown code fences and handles malformed JSON gracefully — LLMs occasionally wrap output in ```` ```json ```` despite explicit instructions not to.

**Prompt-driven domain logic.** The "Finance & Accounting", "Healthcare", "Government" feel of the output comes from prompt context injection, not from training. Adding a new domain or framework is a few lines of code, not a rewrite.

---

## Run it locally

```bash
# Clone the repo
git clone (https://github.com/gankaili88/pitch-deck-generator)
cd pitch-deck-generator

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate    # Windows
# source venv/bin/activate    # Mac/Linux

# Install dependencies
pip install -r requirements.txt

# Get a free Gemini API key from https://aistudio.google.com/apikey
# Create a .env file in the project root with:
# GOOGLE_API_KEY=your-key-here

# Run the app
streamlit run app.py
```

---

## What I learned

- **Structured AI output is hard.** Getting an LLM to reliably return parseable JSON took multiple prompt iterations and defensive parsing. "Reply with only JSON" works 90% of the time; the other 10% needs code-level handling.
- **Prompt engineering > model choice.** The same Gemini 2.5 Flash model produces strikingly different output depending on framework and research context. Most of the project's intelligence lives in the prompt, not in the model.
- **`python-pptx` is unexpectedly powerful.** Building a `.pptx` programmatically — with header strips, footers, page numbers, multi-column layouts — is a few hundred lines of code. Useful workflow for any consultant who automates deck production.
- **Honest limitations are credibility, not weakness.** Documenting what the tool can't do is more useful than overselling what it can.

---

## Limitations (real ones, not boilerplate)

- **Generic competitors when no research context provided.** Without the optional research field filled in, the AI invents plausible-sounding but generic competitor names and numbers.
- **Style selector (McKinsey/BCG/Big Four) is mostly cosmetic.** Output structure varies slightly with style, but not enough to distinguish a "BCG deck" from a "McKinsey deck" in a real test.
- **Niche industries produce confident hallucinations.** Tested on palm oil refining, sovereign wealth fund work, and public-sector education — the tool produces confident-sounding output but with limited domain accuracy. Confident-sounding garbage is the most dangerous failure mode and is documented honestly.
- **No formula or fact verification.** The tool doesn't validate that the numbers it generates are accurate.
- **No persistence.** Each session is independent — no saved decks, no history across sessions.

---

## What's next

- **Slide-by-slide regeneration** — refine one slide without redoing the whole deck (mirrors real consulting iteration with senior managers)
- **Multi-step generation** — "approve the structure, then generate the content" — matches how managers brief associates
- **Logo upload with colour extraction** — currently palettes are picked from a menu; logo-based theming is the natural extension
- **Critique mode** — paste your own deck content, get feedback on weak titles and missing structure

---

## About me

I'm an ICAEW student preparing to join a Big Four firm as a consulting associate. I'm building a series of AI-powered tools to develop fluency in applied LLMs before joining — this is project 2 of an ongoing series.

- 💼 [LinkedIn](https://www.linkedin.com/in/gan-kai-li-a8a782317/)
- 📂 [Excel Formula Generator (Project 1)](https://github.com/gankaili88/excel-formula-generator)
