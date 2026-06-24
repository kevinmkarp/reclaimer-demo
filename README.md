# RE•CLAIMER

RE•CLAIMER is a local Streamlit app for planning whether an AI recommendation agent would treat a marketing claim as decision evidence, brand context, or ignore it, and what would need to change for the claim to affect a recommendation.

It is designed for marketers who want to turn vague or isolated claims into structured, source-validated arguments that AI recommendation agents can understand.

## Run the app

Fastest path in this workspace:

```bash
cd /Users/kevinkarp/Documents/Codex/2026-06-23/files-mentioned-by-the-user-revise/outputs/agent-proof-workbench
chmod +x run_app.sh
./run_app.sh
```

Emergency demo fallback with live web disabled:

```bash
./run_app_no_web.sh
```

Standard Python setup:

```bash
cd /Users/kevinkarp/Documents/Codex/2026-06-23/files-mentioned-by-the-user-revise/outputs/agent-proof-workbench
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Add your OpenAI API key

The app loads the API key in this order:

1. `OPENAI_API_KEY` environment variable
2. Streamlit secrets
3. Sidebar entry as a local fallback

For local development, create a `.env` file in this folder:

```bash
OPENAI_API_KEY=your_api_key_here
```

When a deployment key is configured, judges see only `API key configured for demo.`
The app never hardcodes, prints, logs, displays, or exports the key. `.env` and
`.streamlit/secrets.toml` are included in `.gitignore`.

## Public deployment

### Railway

1. Push this folder to a GitHub repository.
2. In Railway, create a new project and choose **Deploy from GitHub repo**.
3. In the Railway service variables, add:

```text
OPENAI_API_KEY=your_api_key_here
```

4. Deploy the service, then generate a public domain from Railway's networking settings.

Railway uses the included `Procfile` to start Streamlit on its assigned public port.
Do not commit the API key to GitHub.

### Streamlit Community Cloud

Before sharing the public link, add `OPENAI_API_KEY` in the app's
**Secrets / Settings** panel:

```toml
OPENAI_API_KEY = "your_api_key_here"
```

The deployed app will use that secret automatically and will not ask judges to paste a key.

## Research modes

- **No web**: Uses the model's reasoning only. This is the default.
- **Web-grounded**: Allows the model to use web search for category, substantiation, regulatory, and source context. When sources are used, the app shows citations returned by the model.
- **Evidence-only**: Uses uploaded evidence files only and does not use web search.
- **Evidence + web**: Uses uploaded evidence plus web search.

Uploaded files can include `.txt`, `.md`, `.pdf`, `.docx`, and `.csv` files.

## First-page output

The main result page focuses on three cards:

- **Useful vs. Fluff / Needs Clarification**: Breaks the claim into atomic elements, classifies each one, and audits math or logic when numbers are present.
- **Stronger Agent-Ready Argument**: Shows how to rewrite the claim as a structured, comparative, caveated decision argument if the right proof exists.
- **Validation Target Map**: Identifies where proof would need to live online, grouped by owned site, official authority, expert review, commerce editorial, product data, creator demonstration, and community validation.

The planner also includes:

- **Claim Routes**: Generates 2-4 strategically different ways to make a claim recommendation-useful.
- **Recommended Route**: Selects the route with the strongest combination of user-intent fit, proof achievability, and recommendation impact.
- **Agent Source Map: Who Shapes the Answer**: Uses the actual web-search evidence layer to identify discovered and scanned sources, explain what each source must say, and turn gaps into a ranked CMO contact/action list.
- **Source Discovery Summary**: Shows the inferred search intent, queries run, discovered/scanned counts, source mix, and evidence gaps.
- **Source Echo Impact Ladder**: Simulates how every route moves as owned, editorial, review-aggregator, and community evidence accumulates.
- **Intent Match**: Infers the user need or job-to-be-done implied by the claim.
- **General Category Impact vs. Intent-Specific Impact**: Scores whether a claim changes broad recommendations separately from recommendations for the implied user need.
- **Decision Criteria Detected**: Separates primary criteria, secondary criteria, direct and indirect evidence, tradeoffs, and missing proof.
- **Simulated Recommendation**: Gives a plain-English recommendation for the user segment the claim most directly serves.
- **Argument to Seed**: Provides a clean source-ready argument that third parties can validate and repeat in their own voice, plus a separate owned-site citation version.
- **Proof Packet**: Separates the metrics, comparison set, facts, caveats, methodology, currentness, and supporting artifacts needed.
- **Source Echo Simulation**: Shows how owned, editorial, commerce, review-aggregator, community, competitor, and official sources would affect the argument's authority.
- **Third-Party Story Map**: Prioritizes scanned sources, partial support, source targets, outreach targets, community checks, and competitor fact checks.
- **Claim Structure Status**: Estimates whether the claim is merely mention-worthy or has enough structure to affect a recommendation.
- **Current Web Reality**: Shows only sources actually searched, retrieved, or returned by web search in the current run.
- **Validation Target Map**: Lists source targets an AI recommendation agent would likely use to validate this claim type.
- **Counterfactual Source Echo**: Separately simulates what would happen if sources repeated the claim exactly as written versus the recommended upgraded version.
- **Recommended Upgrade**: Converts weak claims into a category-appropriate proof structure, such as price/promo, legal service, performance, technical, durability, or "best for" evidence.

Source status badges distinguish verified sources, target-only sources, hypothetical proof language, user-supplied claims, and sources that were not checked. The app should not describe what a source currently says unless it was actually checked in web-grounded mode.

Third-party attribution is contextual: a source-ready argument does not need to say "according to [source]" when the source itself is the speaker. Explicit attribution belongs in the separate owned-site citation version.

Named sources are not preloaded. In web-grounded modes, discovered and scanned sources must come from the current search evidence layer. Useful sources that were not found or scanned are labeled `Source target only — not found/scanned in this run`.

Web-grounded analysis uses a split pipeline: RE•CLAIMER first generates and runs several short, natural search queries, then passes the verified source bundle into a separate planner call. The full analysis prompt is never used as a web search query.

The live pipeline is token-capped: one compact discovery request uses `gpt-4.1-mini`, the verified source bundle is limited to six sources, and the final `gpt-4.1` planner response is capped to stay below typical TPM limits.

All recommendation scores use one ladder: `Ignored`, `Brand context only`, `Mention-worthy`, `Consideration-worthy`, `Strong consideration`, `Conditional recommendation`, `Recommendation-changing`.

Evidence Coverage uses: `Not checked`, `Low`, `Partial`, `Moderate`, `Strong`, `Robust`. Claim structure is scored independently from source coverage.

The simulator is calibrated by an observed experiment where a vague promotion was repeated but did not change the consideration set, while a structured, comparative, caveated promotion did. That lesson is used as a general principle only; the app does not hardcode any brand, price, competitor, or category logic.

## Important note

This tool is not legal advice. It helps marketers understand what evidence an AI recommendation agent would likely need before treating a claim as usable evidence in a recommendation.
