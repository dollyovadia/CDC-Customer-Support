"""First-name dictionary for replies: Hebrew spelling + gender (F/M).
Keys are lowercase Latin spellings or Hebrew spellings. Extend freely."""
import re, unicodedata

# hebrew name -> gender
HEBREW = {
 # female
 "נועה":"F","נעה":"F","מיכל":"F","יעל":"F","שירה":"F","תמר":"F","מאיה":"F","רונית":"F","רוני":"F","נעמה":"F","הדס":"F","הדסה":"F",
 "אביטל":"F","ענת":"F","אור":"F","הילה":"F","נטע":"F","אילנה":"F","חדווה":"F","חגית":"F","דורין":"F","עדן":"F","אליה":"F","שלומית":"F",
 "ליהי":"F","שני":"F","גאיה":"F","אמנה":"F","שרון":"F","שירלי":"F","ספיר":"F","סיגל":"F","יובל":"F","רלי":"F","ענבל":"F","רחל":"F",
 "ליאל":"F","לינוי":"F","מונירה":"F","מירה":"F","עמית":"F","שונית":"F","תם":"F","גילה":"F","שילה":"F","נורית":"F","אדל":"F","שירן":"F",
 "שיר":"F","עדי":"F","איילת":"F","יפעת":"F","אושר":"F","דניאל":"M","אלישיה":"F","רחלי":"F","צופיה":"F","ניגינה":"F","רותם":"F",
 "סיגלת":"F","מלכה":"F","אניצ'קה":"F","אתי":"F","גאולה":"F","לנה":"F","ליטל":"F","דבורה":"F","רויטל":"F","מירב":"F","ורד":"F",
 "גלית":"F","אנג'לה":"F","ויויאן":"F","לבנת":"F","איריס":"F","ענבר":"F","פז":"F","גל":"F","קרן":"F","נטלי":"F","דנה":"F","צאלה":"F",
 "בילי":"F","אורית":"F","אורלי":"F","אסתר":"F","אפרת":"F","בת-אל":"F","בתאל":"F","גלי":"F","דפנה":"F","דקלה":"F","הודיה":"F",
 "זהבה":"F","חן":"F","חנה":"F","טל":"F","טליה":"F","יהודית":"F","יסמין":"F","ירדן":"F","כרמית":"F","כרמל":"F","לאה":"F","ליאור":"M",
 "ליאת":"F","לימור":"F","לירון":"F","לירז":"F","מור":"F","מוריה":"F","מורן":"F","מיטל":"F","מיכאלה":"F","מירי":"F","מרים":"F",
 "נגה":"F","נוגה":"F","נופר":"F","ניצן":"F","נירית":"F","סיון":"F","סמדר":"F","עינב":"F","עינת":"F","עליזה":"F","עמליה":"F","פנינה":"F",
 "קורל":"F","קרין":"F","רביד":"F","רבקה":"F","רוית":"F","רוזי":"F","רינת":"F","ריקי":"F","שגית":"F","שולמית":"F","שולה":"F","שושי":"F",
 "שלי":"F","שקד":"F","שרה":"F","שרית":"F","תהילה":"F","תמי":"F","תמרה":"F","אביב":"M","אביגיל":"F","אביה":"F","אגם":"F","אודליה":"F",
 "אוריה":"F","אורנה":"F","אלה":"F","אלונה":"F","אליענה":"F","אלינור":"F","אמילי":"F","אנה":"F","אנאל":"F","אסנת":"F","אריאל":"M",
 "בר":"F","ברכה":"F","גבריאלה":"F","דור":"M","דיאנה":"F","דינה":"F","הגר":"F","הדר":"F","הילי":"F","הלן":"F","ויקי":"F","ויקטוריה":"F",
 "זיו":"M","חוה":"F","חיה":"F","יהל":"F","יוליה":"F","יונית":"F","יוכבד":"F","ילנה":"F","יערה":"F","כנרת":"F","לי":"F","ליבי":"F",
 "לידור":"F","ליז":"F","לילך":"F","לינור":"F","מאי":"F","מגי":"F","מורל":"F","מיה":"F","מילה":"F","מעיין":"F","מרינה":"F","משי":"F",
 "נאוה":"F","נבו":"M","נורה":"F","נילי":"F","נינה":"F","נעמי":"F","נתלי":"F","סופי":"F","סימה":"F","סיגלית":"F","סתיו":"F","עדינה":"F",
 "עטרה":"F","עלמה":"F","ענבר":"F","עפרה":"F","פאני":"F","פלורה":"F","צביה":"F","ציפי":"F","ציפורה":"F","קטיה":"F","קרולין":"F",
 "רוחמה":"F","רומי":"F","רות":"F","רותי":"F","ריטה":"F","רננה":"F","רעות":"F","שגב":"M","שהם":"F","שובל":"F","שחר":"F","שלומית":"F",
 "שמרית":"F","שרונה":"F","תאיר":"F","תהל":"F","תומר":"M","תכלת":"F","אלינה":"F","יהודה":"M","מאשה":"F","נטע-לי":"F","נטלי":"F",
 # male
 "ליאור":"M","אורי":"M","יונתן":"M","יהונתן":"M","רון":"M","שי":"M","אופיר":"M","רפאל":"M","חיים":"M","חגי":"M","עומר":"M","אלון":"M",
 "אמיר":"M","אסף":"M","אבי":"M","אברהם":"M","אהרון":"M","איתי":"M","איתמר":"M","אלי":"M","אליהו":"M","אלעד":"M","אלכס":"M","אלכסנדר":"M",
 "אריק":"M","בועז":"M","בן":"M","בני":"M","בנימין":"M","ברק":"M","גבי":"M","גדי":"M","גיא":"M","גיל":"M","גלעד":"M","דוד":"M","דודו":"M",
 "דורון":"M","דן":"M","דני":"M","דרור":"M","הראל":"M","הרצל":"M","ויקטור":"M","זאב":"M","חנן":"M","טום":"M","יאיר":"M","יגאל":"M",
 "יובל":"F","יוסי":"M","יוסף":"M","יורם":"M","יוחאי":"M","יניב":"M","יעקב":"M","יצחק":"M","ירון":"M","ישי":"M","ישראל":"M","כפיר":"M",
 "לירן":"M","מאור":"M","מוטי":"M","מיכאל":"M","מנחם":"M","מני":"M","משה":"M","נדב":"M","נועם":"M","נח":"M","נחום":"M","ניר":"M","ניסים":"M",
 "ניצן":"F","נתן":"M","נתנאל":"M","סער":"M","עדו":"M","עידו":"M","עופר":"M","עוז":"M","עמוס":"M","עמי":"M","עמרי":"M","ערן":"M","פיני":"M",
 "צבי":"M","צחי":"M","קובי":"M","רועי":"M","רז":"M","רמי":"M","רן":"M","רני":"M","רפי":"M","שאול":"M","שגיא":"M","שון":"M","שלומי":"M",
 "שלמה":"M","שמואל":"M","שמעון":"M","תום":"M","אוהד":"M","איל":"M","אייל":"M","אופק":"M","עמית":"F","עדי":"F","אליאס":"M","דמיטרי":"M",
 "יהודה":"M","בועז":"M","נמרוד":"M","ארז":"M","אדם":"M","אביעד":"M","אורן":"M","ארי":"M","אריאל":"M","אלירן":"M","אלמוג":"M","יוני":"M",
 "ליעד":"M","מתן":"M","נבו":"M","ניב":"M","שחר":"F","שקד":"F","טל":"F","עדן":"F","תומר":"M","אביב":"M","דור":"M","זיו":"M","יהל":"F",
}

# latin spelling -> hebrew
LATIN = {
 "nathalie":"נטלי","natalie":"נטלי","natali":"נטלי","keren":"קרן","dana":"דנה","hila":"הילה","neta":"נטע","ilana":"אילנה","liran":"לירן",
 "ron":"רון","shay":"שי","shai":"שי","ofir":"אופיר","eden":"עדן","eliya":"אליה","lihi":"ליהי","shani":"שני","gaia":"גאיה","sharon":"שרון",
 "sapir":"ספיר","haim":"חיים","chaim":"חיים","yael":"יעל","reli":"רלי","shira":"שירה","inbal":"ענבל","rachel":"רחל","liel":"ליאל",
 "linoy":"לינוי","amit":"עמית","shunit":"שונית","shir":"שיר","sigal":"סיגל","ifat":"יפעת","yifat":"יפעת","adi":"עדי","alicia":"אלישיה",
 "elias":"אליאס","nigina":"ניגינה","hanan":"חנן","yossi":"יוסי","ofek":"אופק","anichka":"אניצ'קה","maya":"מאיה","dmitry":"דמיטרי",
 "noa":"נועה","lital":"ליטל","deborah":"דבורה","debora":"דבורה","eyal":"אייל","ohad":"אוהד","shiran":"שירן","galit":"גלית","angela":"אנג'לה",
 "ofer":"עופר","livnat":"לבנת","yuval":"יובל","roni":"רוני","iris":"איריס","inbar":"ענבר","michal":"מיכל","tamar":"תמר","ronit":"רונית",
 "naama":"נעמה","hadas":"הדס","avital":"אביטל","anat":"ענת","or":"אור","hadassa":"הדסה","hagit":"חגית","dorin":"דורין","shlomit":"שלומית",
 "shirly":"שירלי","shirley":"שירלי","mira":"מירה","gila":"גילה","nurit":"נורית","adel":"אדל","ayelet":"איילת","osher":"אושר","daniel":"דניאל",
 "danielle":"דניאל","racheli":"רחלי","tzofia":"צופיה","rotem":"רותם","malka":"מלכה","etti":"אתי","eti":"אתי","geula":"גאולה","lena":"לנה",
 "revital":"רויטל","merav":"מירב","meirav":"מירב","vered":"ורד","vivian":"ויויאן","paz":"פז","gal":"גל","tzeela":"צאלה","billy":"בילי",
 "orit":"אורית","orly":"אורלי","esther":"אסתר","efrat":"אפרת","batel":"בת-אל","gali":"גלי","dafna":"דפנה","dikla":"דקלה","hodaya":"הודיה",
 "chen":"חן","hen":"חן","hana":"חנה","tal":"טל","talia":"טליה","yasmin":"יסמין","yarden":"ירדן","carmel":"כרמל","lea":"לאה","leah":"לאה",
 "liat":"ליאת","limor":"לימור","liron":"לירון","liraz":"לירז","mor":"מור","moriya":"מוריה","moran":"מורן","meital":"מיטל","miri":"מירי",
 "miriam":"מרים","noga":"נוגה","nofar":"נופר","nitzan":"ניצן","sivan":"סיון","smadar":"סמדר","einav":"עינב","einat":"עינת","coral":"קורל",
 "karin":"קרין","ravid":"רביד","rivka":"רבקה","ravit":"רוית","rinat":"רינת","riki":"ריקי","sagit":"שגית","shelly":"שלי","shaked":"שקד",
 "sara":"שרה","sarah":"שרה","sarit":"שרית","tehila":"תהילה","tami":"תמי","aviv":"אביב","avigail":"אביגיל","agam":"אגם","orna":"אורנה",
 "ella":"אלה","alona":"אלונה","elinor":"אלינור","emily":"אמילי","anna":"אנה","osnat":"אסנת","ariel":"אריאל","bar":"בר","diana":"דיאנה",
 "dina":"דינה","hagar":"הגר","hadar":"הדר","vicky":"ויקי","victoria":"ויקטוריה","ziv":"זיו","julia":"יוליה","yulia":"יוליה","yonit":"יונית",
 "yelena":"ילנה","elena":"ילנה","yaara":"יערה","kineret":"כנרת","lee":"לי","li":"לי","libi":"ליבי","lidor":"לידור","liz":"ליז","lilach":"לילך",
 "linor":"לינור","may":"מאי","mia":"מיה","maayan":"מעיין","marina":"מרינה","nava":"נאוה","nili":"נילי","nina":"נינה","naomi":"נעמי",
 "sophie":"סופי","sima":"סימה","stav":"סתיו","alma":"עלמה","ofra":"עפרה","tzipi":"ציפי","katya":"קטיה","romi":"רומי","ruth":"רות","ruti":"רותי",
 "rita":"ריטה","reut":"רעות","shoham":"שהם","shoval":"שובל","shahar":"שחר","tair":"תאיר","tomer":"תומר","alina":"אלינה","masha":"מאשה",
 "lior":"ליאור","uri":"אורי","yonatan":"יונתן","jonathan":"יונתן","raphael":"רפאל","rafael":"רפאל","hagai":"חגי","omer":"עומר","alon":"אלון",
 "amir":"אמיר","asaf":"אסף","assaf":"אסף","avi":"אבי","avraham":"אברהם","itay":"איתי","itai":"איתי","itamar":"איתמר","eli":"אלי","elad":"אלעד",
 "alex":"אלכס","arik":"אריק","boaz":"בועז","ben":"בן","benny":"בני","barak":"ברק","gabi":"גבי","gadi":"גדי","guy":"גיא","gil":"גיל",
 "gilad":"גלעד","david":"דוד","dudu":"דודו","doron":"דורון","dan":"דן","danny":"דני","dror":"דרור","harel":"הראל","victor":"ויקטור",
 "yair":"יאיר","yosef":"יוסף","yoram":"יורם","yaniv":"יניv","yaakov":"יעקב","yitzhak":"יצחק","yaron":"ירון","yishai":"ישי","kfir":"כפיר",
 "maor":"מאור","moti":"מוטי","michael":"מיכאל","menachem":"מנחם","moshe":"משה","nadav":"נדב","noam":"נועם","nir":"ניר","natan":"נתן",
 "netanel":"נתנאל","saar":"סער","ido":"עידו","oz":"עוז","amos":"עמוס","omri":"עמרי","eran":"ערן","tzvi":"צבי","kobi":"קובי","roy":"רועי",
 "roee":"רועי","raz":"רז","rami":"רמי","ran":"רן","rafi":"רפי","shaul":"שאול","sagi":"שגיא","shlomi":"שלומי","shlomo":"שלמה","shmuel":"שמואל",
 "tom":"תום","eliran":"אלירן","almog":"אלמוג","yoni":"יוני","matan":"מתן","niv":"ניב","nimrod":"נמרוד","erez":"ארז","adam":"אדם","oren":"אורן",
 "yehuda":"יהודה","dor":"דור","nevo":"נבו","yahel":"יהל","oshri":"אושרי","mazal":"מזל","batya":"בתיה","yocheved":"יוכבד","zehava":"זהבה",
}
LATIN["yaniv"] = "יניב"
HEBREW.update({'אביבה': 'F', 'אדווה': 'F', 'אורטל': 'F', 'אסתי': 'F', 'בלהה': 'F', 'גיולנרה': 'F', 'דורית': 'F', 'חופני': 'M', 'חנית': 'F', 'טוהר': 'F', 'יפה': 'F', 'ירדנה': 'F', 'כרמלה': 'F', 'ליבנת': 'F', 'מוניק': 'F', 'סופיה': 'F', 'עדנה': 'F', 'עידית': 'F', 'עידן': 'M', 'שי-לי': 'F', 'שילת': 'F', 'אחמד': 'M', "אלינצ'יק": 'F', 'עליזה': 'F', 'ארתור': 'M', 'אביחי': 'M', 'אבישי': 'M', 'בטי': 'F', 'קמיל': 'F', 'אלעזר': 'M', 'ינון': 'M', 'עירית': 'F', "ג'סיקה": 'F', "ג'וני": 'M', "ח'דיג'ה": 'F', 'לילה': 'F', 'אופירה': 'F', 'פזית': 'F', 'פזי': 'F', 'רביב': 'M', 'רותה': 'F', 'שאדי': 'M', 'שריפה': 'F', 'שיראל': 'F', 'טלי': 'F', 'טטיאנה': 'F', 'צביה': 'F', 'חופית': 'F', 'ראיסה': 'F', 'נוי': 'F', 'יעקב': 'M', 'רומי': 'F', 'מוריה': 'F', 'אורלי': 'F'})
LATIN.update({'ahmad': 'אחמד', 'alinchik': "אלינצ'יק", 'aliza': 'עליזה', 'amneh': 'אמנה', 'artur': 'ארתור', 'avihai': 'אביחי', 'avishai': 'אבישי', 'betti': 'בטי', 'camil': 'קמיל', 'elazar': 'אלעזר', 'etty': 'אתי', 'inon': 'ינון', 'irit': 'עירית', 'jessica': "ג'סיקה", 'johnny': "ג'וני", 'khadeja': "ח'דיג'ה", 'leila': 'לילה', 'moria': 'מוריה', 'ofira': 'אופירה', 'orli': 'אורלי', 'pazit': 'פזית', 'pazy': 'פזי', 'raviv': 'רביב', 'rommy': 'רומי', 'rutha': 'רותה', 'shadi': 'שאדי', 'sharifa': 'שריפה', 'shirel': 'שיראל', 'sofia': 'סופיה', 'tali': 'טלי', 'tatiana': 'טטיאנה', 'yakov': 'יעקב', 'zvia': 'צביה', 'hofit': 'חופית', 'раиса': 'ראיסה', 'noy': 'נוי'})


def clean_first_name(raw):
    """'💫 Eden Barsano 💫' -> 'eden';  'ענת כהן אל גלי' -> 'ענת'."""
    if not raw: return ""
    s = unicodedata.normalize("NFKC", str(raw))
    s = s.replace("_", " ").replace(".", " ")
    s = "".join(ch for ch in s if unicodedata.category(ch)[0] in "LZ" or ch in "'-").strip()  # letters, spaces, ' and -
    return s.split()[0].lower() if s.split() else ""


def hebrew_name_and_gender(raw):
    """Return (hebrew_first_name, gender F/M/'' , needs_check bool)."""
    first = clean_first_name(raw)
    if not first: return "", "", True
    if re.search(r"[֐-׿]", first):
        return first, HEBREW.get(first, ""), first not in HEBREW
    heb = LATIN.get(first)
    if heb: return heb, HEBREW.get(heb, ""), heb not in HEBREW
    return "", "", True   # unknown latin name: leave the name out of the greeting


# ---------------------------------------------------------- name <-> email
_HEB_SKEL = {"ב":"b","ג":"g","ד":"d","ז":"z","ח":"h","ט":"t","כ":"k","ך":"k","ל":"l","מ":"m","ם":"m","נ":"n","ן":"n","ס":"s",
             "פ":"p","ף":"p","צ":"ts","ץ":"ts","ק":"k","ר":"r","ש":"s","ת":"t"}   # א ה ו י ע dropped (vowel-like)
_LAT_TO_HEB_CANON = {}
for _lat, _heb in LATIN.items():
    _LAT_TO_HEB_CANON.setdefault(_heb, _lat)


def skeleton(s):
    """Consonant skeleton so Hebrew and Latin spellings compare: נטלי -> ntl, nathalie -> ntl."""
    s = unicodedata.normalize("NFKC", str(s or "")).lower()
    s = s.replace("ph", "f").replace("sh", "s").replace("ch", "h").replace("tz", "ts").replace("ck", "k")
    out = []
    for ch in s:
        if ch in _HEB_SKEL: out.append(_HEB_SKEL[ch])
        elif "֐" <= ch <= "׿": continue
        elif ch in "aeiouyhw'-": continue
        elif ch == "v": out.append("b")
        elif ch in "cq": out.append("k")
        elif ch == "f": out.append("p")
        elif ch.isalpha(): out.append(ch)
    return "".join(out)


def name_tokens(raw):
    s = unicodedata.normalize("NFKC", str(raw or "")).replace("_", " ").replace(".", " ").replace("-", " ")
    s = "".join(ch for ch in s if unicodedata.category(ch)[0] in "LZ" or ch == "'")
    toks = [t.lower() for t in s.split() if len(t) >= 2]
    extra = []
    for t in toks:
        if t in _LAT_TO_HEB_CANON: extra.append(_LAT_TO_HEB_CANON[t])   # hebrew first name -> its latin spelling
        if t in LATIN: extra.append(LATIN[t])
    return toks + extra


def email_matches_name(email, customer_name):
    """True when the email's local part contains something that looks like the
    customer's first name or surname (skeleton match, 3+ consonants, or the
    full name when shorter)."""
    local = (email or "").split("@")[0].lower()
    if not local or not customer_name: return False
    ls = skeleton(local)
    for tok in name_tokens(customer_name):
        ts = skeleton(tok)
        if len(ts) >= 3 and ts in ls: return True
        if 1 <= len(ts) < 3 and (tok in local or (ts and ls.startswith(ts))): return True
    return False


def mask_email(email):
    if not email or "@" not in email: return email
    local, dom = email.split("@", 1)
    return f"{local[:2]}***@{dom}"
