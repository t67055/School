"""Keyword codebooks for tagging open-ended survey answers (multi-label)."""
import re

NONE_PAT = r"no major concerns?|^no concerns?|nothing comes to mind|none at present|i.?m not sure|i literally have no clue|^unkno?w|^not really\W*$|^\W*(n/?a|none|nothing|no idea|idk|i ?don.?t know|not sure|unsure|no comment|same|-+|\?+|/)\W*$|^i don.?t know enough|^nothing i can think of"

GENERAL = [
 ('Access, cost & equity', r"topic (2|ii)\b|^\W*\(?(ii|2)\)?\W*$|bills|greed|access|afford|\bcosts?\b|costly|expens|pric(e|es|ing)\b|insurance|medicare|medicaid|equit|inequal|unequal|dispar|wealth|\brich\b|\bpoor\b|\bgap\b|democrati|everyone|for all\b|to all\b|all people|fair|reach of|only benefit|well off|socioeconomic|universal health"),
 ('Healthcare system & caregiving', r"care ?tak|take care of (old|elder|them|older)|support(ing)? (for )?(the )?(older|elder)|lack of support|elderly care|long ?term care|facilit|services taking care|medical (attention|equipment|pressure)|people to take care|health ?care|medical (care|system|resources|support|services)|doctor|hospital|nurs(e|es|ing)|caregiv|care ?home|long-term care|assisted living|elder ?care|clinic|physician|practitioner|providers|care facilit|someone to (help|care)|care for them|home care|geriatric"),
 ('Chronic disease & physical health', r"injur|energy|exercis|being active|unhealth|healthiness|healthy|erectile|chemicals in|chronic|disease|cancer|heart|cardio|physical|mobility|frail|\borgans?\b|immun|\bfalls?\b|pain|bod(y|ily)|strength|disabilit|illness|diabet|wound|regenerat|(?<!mental )(?<!reproductive )\bhealth\b(?!\s*care)(?! insurance)(?! effects)|sick|stroke|muscle|bone"),
 ('Cognitive decline & dementia', r"dementia|alzheimer|cogniti|memory|\bbrain|neuro"),
 ('Mental & emotional health', r"\bmental|emotional|depress|anxiety|psycholog|well-?being|\bstress"),
 ('Social connection & isolation', r"lonel|lonli|isolat|social(?! security)|community|friend|famil|relationship|connect|companion|support system"),
 ('Financial security, work & basic needs', r"housing|basic (needs|living)|welfare|survival|living security|cost of living|consumption|financ|money|income|retire|pension|social security|saving|inflation|\bjobs?\b|employ|work longer|keep working|paying work|poverty|economic|fixed income|\brent\b"),
 ('Independence & autonomy', r"independen|autonom|self-suffic|own decisions|daily (tasks|life|activities|skills)|everyday tasks|take care of (themselves|myself)"),
 ('Technology & AI', r"topic (3|iii)\b|^\W*\(?(iii|3)\)?\W*$|number three|third (topic|option)|\ba1\b|ux\b|\bai\b|a\.i\.|artificial intelligence|machine learning|\bml\b|technolog|algorithm|digital|\btech\b|computer|robot|\bmodels?\b|large datasets?|big data"),
 ('Privacy & data protection', r"privacy|private|confidential|surveillance|personal (data|information)|data (protection|security|usage|use|ownership)|health data|data privacy|ownership"),
 ('Safety, accuracy & misuse', r"distort|viable|rogue|mislead|exploit|negative way|\bsafe|safety|risk|harm|side effect|unintended|unknown|danger|accura|error|wrong|mistake|misinformation|false|reliab|over ?-?relian|abuse|misuse|bioweapon|ill intention|too far|playing god"),
 ('Ethics, regulation & policy', r"leadership|administration|funding|humanity|human values|human beings|statistics|assisted suicide|cloning|surrogacy|concent|ethic|regulat|oversight|guardrail|polic(y|ies)|government|\blaws?\b|legal|legislat|transparen|accountab|\bbias|discrimin|eugenic|consent|politic|congress|governance|moral|rules|guidelines"),
 ('Research, science & treatments', r"progress|medical techniques|glp-?1|solutions for aging|modern solutions|research|scien|innovat|breakthrough|discover|\bcures?\b|treatment|therap|\bdrugs?\b|medicin|advance|clinical trial|stud(y|ies)|experiment|testing|engineer"),
 ('Prevention & early detection', r"prevent|\bearly|detect|screen|diagnos|predict|monitor|biomarker|precision|personali[sz]"),
 ('Longevity & lifespan', r"longevity|lifespan|life span|live longer|living longer|life extension|extend(ing)? (life|lifespan|human)|healthspan|longer lives|longer life|live healthier|healthier for longer|life expectancy|anti-?aging"),
 ('Reproductive health & fertility', r"topic (1|i)\b|^\W*\(?(i|1)\)?\W*$|reproduct|fertil|pregnan|children|\bivf\b|menopaus|\bbirth|\beggs?\b|sperm|\bkids\b|older parents|having a baby"),
 ('Ageism & how society treats older adults', r"ageis|stigma|respect|perception|attitude|treated as|ignored|overlooked|marginali|left behind|understand them|people too|valued|appreciated|dignity|\bproductive|placement in society|integral part"),
 ('Population & societal burden', r"young people|birth ?rate|new ?born|babies|childbearing|workforce|dependency ratio|national development|creativity in society|population|demograph|aging society|burden|overpopulat|resources|sustainab|younger generation|infrastructure|climate|environment|water"),
 ('Quality of life', r"quality of life|happi|enjoy|fulfil|purpose|meaning|comfort|live well|better lives|wellness|good life"),
]

DEFINITION = [
 ('Mindset / "just a number"', r"just a number|blessing|mindset|attitude|state of mind|how you feel|scary|frighten|painful"),
 ('Decline & change in function', r"decreas|unhealth|entropy|function|independen|resilien|declin|deteriorat|degenerat|decay|wear|weak|loss|lose|losing|worse|slow|less |reduc|difficult|vulnerab|break(s|ing)? down|impair|deplet|frail"),
 ('Biological / cellular process', r"biolog|\bcells?\b|cellular|telomer|senescen|\bdna\b|genom|genetic|molecular|physiolog|homeostasis|\borgans?\b|tissue|metabol"),
 ('Passage of time / getting older', r"living longer|longer|alive|born\b|lifelong|life ?span|\d0s\b|mid ?point|over ?\d\d|\btime\b|years|older|chronolog|birth|clock|\bage\b|getting old|growing old|passage|over70|existing"),
 ('Physical & bodily change', r"fitness|\blook|physical|bod(y|ily)|appearance|wrinkl|mobility|hair|teeth|tooth|energy|recover|strength"),
 ('Cognitive & mental change', r"\bmental|\bmind|cogniti|memory|brain|psycholog|emotional"),
 ('Growth, experience & wisdom', r"grow(th|ing)? (and|through|wiser)|wis(e|dom)|experience|matur|learn|journey|positive|wiser|develop"),
 ('Natural stage of life', r"natural|inevitab|stage of life|life cycle|circle of life|part of life|normal|life process|process of life"),
 ('Social & role change', r"social|role|retire|society|relationship"),
]

MAINTAIN = [
 ('Physical health & mobility', r"sleep|\beat|nutri|\brest\b|jump|stand up|physical|(?<!mental )(?<!mental/cognitive )(?<!cognitive )\bhealth\b|mobility|\bbody|movement|move\b|moving|fitness|\bfit\b|strength|athlet|energy|vision|eyesight|hearing|pain|active|walk|exercise|sport|stamina|flexib|balance"),
 ('Memory & cognitive sharpness', r"imagination|critical|open-?minded|logic|memor|cogniti|\bmind\b|sharp|\bbrain|thinking|intellig|reasoning|clarity|lucid|faculties|learning ability|mental (acuity|capacit|abilit|facult|function)|wits|intellect|mentally facilities|cognition"),
 ('Mental & emotional well-being', r"emotional|mental health|mental/cognitive|well-?being|happi|happy|joy|peace|content|mood|positiv|optimis|humou?r|sanity|mental stability|hope|spark|youthful"),
 ('Relationships & social life', r"famil|friend|relationship|social|loved|connection|community|partner|spouse|husband|wife|kids|children|people"),
 ('Independence', r"independ|autonom|take care of myself|own decisions|self-suffic|care for myself|on my own"),
 ('Financial security', r"financ|money|wealth|stability|income|career|\bjob|home\b|house|saving"),
 ('Hobbies, curiosity & purpose', r"hobb|travel|\bread|learn|curio|passion|purpose|creativ|interest|adventure|outdoor|enjoy|music|sport|ambition|drive|goals|work\b|knowledge|lane"),
 ('Values & identity', r"conviction|intentions|be good|do right|faith|values|integrity|kindness|\bkind\b(?! of)|compassion|empathy|discipline|personality|identity|dignity|reputation|spirit|religio|morals"),
 ('Appearance', r"appearance|looks|hair|skin|teeth|beauty|youthful look|cholesterol"),
]

ACTIONS = [
 ('Healthy lifestyle & habits (general)', r"lifestyle|habit|routine|consisten|disciplin|lock in|take care of myself|taking care of myself|stay(ing)? healthy|healthy|\bfit\b|keep (doing|practicing)|maintain (my|all)|trying to"),
 ('Not doing anything yet', r"not currently|not at the moment|nothing i am|^no\b|no guaranteed|can.?t predict|not much i can|plan to\b|need to start"),
 ('Exercise & physical activity', r"excers|weights|exercis|work ?out|gym|active|physical activity|walk|\brun|sport|yoga|move|fitness|lift|swim|hik|stretch|sedentary|\btrain"),
 ('Diet & nutrition', r"diet|\beat|nutri|food|cook|vegetable|healthy eating|supplement|vitamin"),
 ('Sleep & rest', r"sleep|\brest\b"),
 ('Avoiding harmful habits', r"smok|alcohol|drink|substance|bad habits|avoid|antihistamine"),
 ('Social engagement', r"social|friend|famil|connect|relationship|community|people|talk|volunteer"),
 ('Learning & mental stimulation', r"learn|\bread|puzzle|brain|mental(ly)? (exercise|active|stimul|acuity)|cogniti|educat|skill|curio|journal|games|stay engaged|engag|mentally"),
 ('Preventive & medical care', r"check|screen|doctor|medical|prevent|health ?care|medicine|vaccin|physician|dentist|follow medical"),
 ('Stress & emotional well-being', r"stress|mindful|meditat|mood|positiv|happy|balance|peace|gratitude|mental health|emotional|calm|grounded|relax|why\b"),
 ('Financial planning', r"financ|sav(e|ing)|money|invest|retire|plan(ning)? (ahead|financially)|plan ahead|insurance|budget"),
]

VAGUE_PAT = r"^\W*(yes|yeah|yep|always|sure|of course|definitely|most definitely|absolutely|certainly|for sure|maybe|perhaps|probably)\b.{0,60}$"

REASON_EXTRA = [
 ('Personal relevance', r"personal|to me\b|for me\b|my (own|life|family|needs|health|future|life stage)|myself|individual(ly)?\b|relevant to me|i care|i want|i would like|important to me|at my age|my age|\bme\b"),
 ('Broad benefit to society', r"most people|largest number|greatest number|majority|benefit|impact|everyone|society|public|population|widespread|more people|all people|humanity"),
 ('All connected / equally important', r"equally|all (related|connected|important)|hand in hand|interconnect|depend on each other|without the (other|last|first)|build(s)? on|foundation"),
]

def tag(text, codebook):
    if not text:
        return []
    t = str(text).lower().strip()
    tags = [name for name, pat in codebook if re.search(pat, t)]
    if not tags:
        if re.search(NONE_PAT, t):
            return ['No answer / unsure']
        if re.search(VAGUE_PAT, t):
            return ['General / no specifics']
        return ['Other']
    return tags

def action_status(text):
    if not text:
        return None
    t = str(text).lower().strip()
    if re.match(r"^(yes,? substantial|yes\b|yeah|definitely|absolutely|it has been addressed)", t):
        return 'Yes'
    if re.search(r"not particularly|has not|not in this|no action|do not see|little or no|very limited|^no\b|^nope|not really|not much|^none|not seeing|i don.?t see|haven.?t seen|cannot recall|can.?t think|not that i|nothing", t):
        return 'Little or none'
    if re.search(r"^some|somewhat|partially|limited|a little|insufficient|not enough|slow|kind of|starting|some action|to some extent|^a bit", t):
        return 'Some, but insufficient'
    if re.search(r"do not know|not sure|^idk|don.?t know|unsure|^maybe|not aware|i.?m not a", t):
        return 'Unsure'
    return 'Names an action or idea'
