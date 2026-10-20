# Aging Survey Explorer

Project 1 surveyed 732 people across 8 class survey groups about aging, longevity research and the role of bioengineers. Each group ran its own version of the survey, so this folder holds the code that merges them into one consistent dataset, tags the open-ended answers by theme, and an interactive page for comparing answers across demographics.

## The site

`index.html` is a self-contained page (it loads d3 from cdnjs and reads `data/survey_responses.json`). With GitHub Pages turned on for this repo (Settings → Pages → Deploy from branch → `main`, folder `/ (root)`), it is served at:

`https://t67055.github.io/School/aging-survey-explorer/`

To view it locally, serve the folder rather than opening the file directly, since browsers block `fetch` from `file://`:

```bash
cd aging-survey-explorer
python3 -m http.server 8000   # then open http://localhost:8000
```

What it shows, with one dot per group for whichever comparison you pick:

- **Respondents:** all groups, or only Group 2.
- **Filter by location:** where people live or where they're from, by country and then state or province. Places with 10+ respondents are listed, plus Greece.
- **Compare by:** age group (Student age (≤29), Middle-aged (30–64), Elderly (65+)), age by decade, age × gender (6 groups), gender, where they live, where they're from, origin & residence, survey group, Group 2 vs rest.
- Topic ranked #1 (personal and societal), yes/no attitudes and 1–5 ratings, guesses of the US 65+ share (actual: 18.0% in 2024, [US Census Bureau](https://www.census.gov/newsroom/press-releases/2025/older-adults-outnumber-children.html)), most exciting / concerning topic.
- Themes in ten open-ended questions, the largest differences between groups, and the answers behind any theme.

Hollow dots mark groups with fewer than 20 answers.

## Data

- `data/Project_1_combined_dataset.xlsx`: the combined workbook.
  - **Combined:** one row per respondent, consistent columns. Locations are split into state/province and country, and age categories are formulas driven by editable cutoffs on the **Notes** sheet.
  - **Tagged Responses:** each open answer beside its theme tags.
  - **Theme Summary:** formula-driven theme shares by age, gender and residence.
  - The original eight group sheets, kept unchanged.
  - The group roster and interviewer names are removed from this public copy.
- `data/survey_responses.json`: one row per respondent, as read by the site.

## Rebuilding

```bash
pip install -r requirements.txt
cd scripts
python3 build.py <original_survey.xlsx> combined.xlsx        # merge + standardize the 8 group sheets
python3 add_tags.py combined.xlsx tagged.xlsx                 # tag open answers, add Tagged Responses + Theme Summary
python3 export_site_data.py ../data/survey_responses.json     # data file for the site
python3 make_public_copy.py tagged.xlsx ../data/Project_1_combined_dataset.xlsx   # drop roster + interviewer names
```

`build.py` and `add_tags.py` write formulas, so open the workbook in Excel or recalculate it in LibreOffice to see the values.

## Method notes

- Answers that had spilled onto extra rows in the original sheets are folded back into the respondent above.
- Ranks are 1 = most important for every group. Groups 2 and 5 typed an ordered list, which is read as topics from most to least important.
- Yes/No answers are taken from the answer's opening word; longer answers are kept as comments.
- Themes come from keyword codebooks in `scripts/tagger.py`: one shared set for challenges, priorities, concerns, excitement and ethics, and separate sets for the definition of aging, things to maintain, and personal actions. An answer can carry several themes. Tagging is approximate, so spot-check the Tagged Responses sheet before quoting a figure.
