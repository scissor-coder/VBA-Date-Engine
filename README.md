# ABID__DateENGINE v2.0 — Complete Documentation
# আবিদ__তারিখইঞ্জিন v2.0 — সম্পূর্ণ ডকুমেন্টেশন

> **A powerful Microsoft Word VBA macro engine for simultaneous Bengali, Gregorian, and Hijri date conversion — with tri-lingual output in Bengali, English, and Arabic.**
>
> **একটি শক্তিশালী Microsoft Word VBA ম্যাক্রো ইঞ্জিন যা বাংলা, গ্রেগরিয়ান এবং হিজরি তারিখ একসাথে রূপান্তর করে — বাংলা, ইংরেজি এবং আরবিতে ত্রিভাষিক আউটপুট সহ।**

---

## Table of Contents | সূচিপত্র

| # | English | বাংলা |
|---|---------|-------|
| 1 | [Overview](#1-overview--সংক্ষিপ্ত-বিবরণ) | সংক্ষিপ্ত বিবরণ |
| 2 | [Architecture At a Glance](#2-architecture-at-a-glance--আর্কিটেকচার-একনজরে) | আর্কিটেকচার একনজরে |
| 3 | [Features](#3-features--বৈশিষ্ট্যসমূহ) | বৈশিষ্ট্যসমূহ |
| 4 | [System Requirements](#4-system-requirements--সিস্টেম-প্রয়োজনীয়তা) | সিস্টেম প্রয়োজনীয়তা |
| 5 | [Installation](#5-installation--ইনস্টলেশন) | ইনস্টলেশন |
| 6 | [Resource File Structure](#6-resource-file-structure--রিসোর্স-ফাইল-কাঠামো) | রিসোর্স ফাইল কাঠামো |
| 7 | [Complete Variable Reference](#7-complete-variable-reference--সম্পূর্ণ-ভেরিয়েবল-রেফারেন্স) | সম্পূর্ণ ভেরিয়েবল রেফারেন্স |
| 8 | [Function Reference](#8-function-reference--ফাংশন-রেফারেন্স) | ফাংশন রেফারেন্স |
| 9 | [How the Algorithm Works](#9-how-the-algorithm-works--অ্যালগরিদম-কিভাবে-কাজ-করে) | অ্যালগরিদম কিভাবে কাজ করে |
| 10 | [Usage Examples & Recipes](#10-usage-examples--recipes--ব্যবহারের-উদাহরণ-ও-রেসিপি) | ব্যবহারের উদাহরণ ও রেসিপি |
| 11 | [Top Tips & Best Practices](#11-top-tips--best-practices--শীর্ষ-টিপস-ও-সেরা-অনুশীলন) | শীর্ষ টিপস ও সেরা অনুশীলন |
| 12 | [Performance Considerations](#12-performance-considerations--পারফরম্যান্স-বিবেচনা) | পারফরম্যান্স বিবেচনা |
| 13 | [Troubleshooting Guide](#13-troubleshooting-guide--সমস্যা-সমাধান-গাইড) | সমস্যা সমাধান গাইড |
| 14 | [Developer Notes & Upgrade Roadmap](#14-developer-notes--upgrade-roadmap--ডেভেলপার-নোট-ও-আপগ্রেড-রোডম্যাপ) | ডেভেলপার নোট ও আপগ্রেড রোডম্যাপ |
| 15 | [Version History](#15-version-history--সংস্করণ-ইতিহাস) | সংস্করণ ইতিহাস |
| 16 | [Contributing](#16-contributing--অবদান) | অবদান |
| 17 | [License](#17-license--লাইসেন্স) | লাইসেন্স |
| 18 | [Credits](#18-credits--কৃতিত্ব) | কৃতিত্ব |

---

## 1. Overview | সংক্ষিপ্ত বিবরণ

### English

**ABID__DateENGINE** is a single-call VBA subroutine that takes one Gregorian date string and simultaneously calculates the corresponding **Bengali (Bangabda)**, **Gregorian (AD)**, and **Hijri (AH)** dates. It then writes those dates — in all three scripts — into **63 public global variables** that stay accessible to every other macro in your Word document for as long as the session is open.

Think of it as a date "pre-processor": you call it once at the top of your macro, and from that point on every calendar and every language is already resolved into neat, ready-to-paste variables. No repeated logic, no lookup tables in your own code, no manual suffix arithmetic.

**Where is it useful?**
- Official government letter-heads requiring all three calendar dates
- Legal documents in Bangladesh that need both বঙ্গাব্দ and Hijri dates
- Religious notices combining Islamic dates with Bengali dates
- Academic certificates with Gregorian and Bengali dates side by side
- Any Word automation project targeting a Bengali-speaking audience

### বাংলা

**আবিদ__তারিখইঞ্জিন** একটি একক-কল VBA সাবরুটিন যা একটি গ্রেগরিয়ান তারিখ স্ট্রিং নিয়ে একই সাথে **বাংলা (বঙ্গাব্দ)**, **গ্রেগরিয়ান (খ্রিস্টাব্দ)**, এবং **হিজরি (হিজরি)** তারিখ গণনা করে। তারপর সেই তারিখগুলো — তিনটি লিপিতেই — **৬৩টি পাবলিক গ্লোবাল ভেরিয়েবলে** লিখে দেয়, যা আপনার Word ডকুমেন্টের যেকোনো ম্যাক্রো থেকে সেশন চলাকালীন যেকোনো সময় অ্যাক্সেসযোগ্য।

এটিকে একটি তারিখ "প্রি-প্রসেসর" হিসেবে ভাবুন: আপনি আপনার ম্যাক্রোর শুরুতে একবার কল করেন, এবং সেখান থেকে প্রতিটি ক্যালেন্ডার ও প্রতিটি ভাষা ইতিমধ্যে সুন্দর, সরাসরি-পেস্টযোগ্য ভেরিয়েবলে রেডি হয়ে যায়।

**কোথায় কাজে আসে?**
- তিনটি ক্যালেন্ডার তারিখ প্রয়োজন সরকারি অফিসিয়াল লেটারহেডে
- বাংলাদেশের আইনি নথিতে বঙ্গাব্দ ও হিজরি তারিখ
- ইসলামিক তারিখের সাথে বাংলা তারিখ যুক্ত ধর্মীয় নোটিশে
- পাশাপাশি গ্রেগরিয়ান ও বাংলা তারিখ সহ একাডেমিক সার্টিফিকেটে

---

## 2. Architecture At a Glance | আর্কিটেকচার একনজরে

### English

The engine is intentionally designed around a simple, two-layer architecture:

```
┌─────────────────────────────────────────────────────────┐
│                   YOUR MACRO (caller)                   │
│        Call ABID__DateENGINE("15/04/2024")              │
└───────────────────────┬─────────────────────────────────┘
                        │ one call
                        ▼
┌─────────────────────────────────────────────────────────┐
│              ABID__DateENGINE  (Public Sub)             │
│                                                         │
│  ① Parse DD/MM/YYYY  →  VBA Date object                │
│  ② LoadDateResources()  ←  reads resource .txt (UTF-8) │
│  ③ Gregorian  calculation  (VBA native)                 │
│  ④ Hijri      calculation  (vbCalHijri switch)          │
│  ⑤ Bangla     calculation  (BD Govt. fixed algorithm)  │
│  ⑥ ToLocalizedNum()  ×  all digits  ×  3 scripts       │
│  ⑦ Suffix logic  (লা/রা/ঠা/ই/শে)                      │
│  ⑧ Write → 63 Public global variables                  │
└─────────────────────────────────────────────────────────┘
                        │ variables now populated
                        ▼
┌─────────────────────────────────────────────────────────┐
│              63 PUBLIC GLOBAL VARIABLES                 │
│   BanglaDate ×  (dayNum + monthNum + yearNum +          │
│                  weekName + monthName + Era)  × bn/en/ar│
│   EnglishDate × (same structure)                        │
│   ArabicDate  × (same structure)                        │
└─────────────────────────────────────────────────────────┘
```

**Key design decisions:**
- **Resource-driven**: All locale strings (month names, weekday names, numerals, suffixes) live outside the `.bas` file in a plain UTF-8 text file. Changing a translation never requires touching VBA code.
- **Global variable pattern**: Instead of returning a struct or object (which VBA makes awkward), the engine writes to module-level public variables. Any macro anywhere in the project reads them directly.
- **Late binding**: ADODB.Stream and FSO are created with `CreateObject`, meaning no compile-time references are needed beyond the standard Word VBA environment.

### বাংলা

ইঞ্জিনটি ইচ্ছাকৃতভাবে একটি সহজ, দুই-স্তরীয় আর্কিটেকচারের উপর নির্মিত:

**মূল ডিজাইন সিদ্ধান্তসমূহ:**
- **রিসোর্স-চালিত**: সমস্ত লোকেল স্ট্রিং (মাসের নাম, সপ্তাহের নাম, সংখ্যা, প্রত্যয়) `.bas` ফাইলের বাইরে একটি সাধারণ UTF-8 টেক্সট ফাইলে থাকে। অনুবাদ পরিবর্তন করতে কখনও VBA কোড স্পর্শ করতে হয় না।
- **গ্লোবাল ভেরিয়েবল প্যাটার্ন**: struct বা object ফেরত দেওয়ার বদলে ইঞ্জিন মডিউল-স্তরের পাবলিক ভেরিয়েবলে লেখে। প্রজেক্টের যেকোনো জায়গার যেকোনো ম্যাক্রো সরাসরি পড়তে পারে।
- **লেট বাইন্ডিং**: ADODB.Stream এবং FSO `CreateObject` দিয়ে তৈরি হয়, তাই স্ট্যান্ডার্ড Word VBA পরিবেশের বাইরে কোনো কম্পাইল-টাইম রেফারেন্সের দরকার নেই।

---

## 3. Features | বৈশিষ্ট্যসমূহ

### English

#### ① Triple Calendar Conversion
One call → three calendars. The input date is simultaneously resolved as:

| Calendar | Era | System | Example (for 15 April 2024) |
|----------|-----|--------|-----------------------------|
| Bengali (Bangabda) | বঙ্গাব্দ | Solar, BD Government fixed | ২রা বৈশাখ ১৪৩১ বঙ্গাব্দ |
| Gregorian | AD | International standard solar | April 15, 2024 AD |
| Hijri | AH | Islamic lunar | ৬ই শাওয়াল ১৪৪৫ হিজরি |

#### ② Triple Script Output
Every date component is immediately available in three scripts. You never need to transliterate manually:

| Script | Language Index | Day example | Month example | Year example |
|--------|---------------|-------------|---------------|--------------|
| Bengali | 1 (bn) | ১৫ | বৈশাখ | ১৪৩১ |
| English | 2 (en) | 15 | Baishakh | 1431 |
| Arabic  | 3 (ar) | ١٥ | بيشاخ | ١٤٣١ |

#### ③ Smart Bengali Date Suffix (প্রত্যয়)
Correctly applies ordinal suffixes used in formal Bengali writing:

| Day(s) | Suffix | Written form | Meaning |
|--------|--------|--------------|---------|
| 1 | লা | ১লা | 1st |
| 2, 3 | রা | ২রা, ৩রা | 2nd, 3rd |
| 4 | ঠা | ৪ঠা | 4th |
| 5 – 18 | ই | ৫ই, ১০ই, ১৮ই | 5th – 18th |
| 19 – 31 | শে | ১৯শে, ২৫শে, ৩০শে | 19th – 31st |

This suffix logic is applied independently to **all three calendars** — so you get `dayNum_bn_pre` for BanglaDate, EnglishDate, and ArabicDate.

#### ④ 63 Public Global Variables
The engine writes **63 named public variables** covering every calendar × every component × every script combination. They persist for the entire VBA session and are readable by any other macro in the document.

#### ⑤ External Resource File
All locale data lives in `ABID_DateENGINE_resources.txt`. To add a new language or correct a transliteration, edit the text file — no VBA editing required.

#### ⑥ Leap Year Awareness
Correctly handles the Gregorian leap year boundary that falls inside the Bengali month of Falgun (ফাল্গুন), adding one day to keep the Bengali date accurate on Feb 29.

### বাংলা

#### ① ত্রিগুণ ক্যালেন্ডার রূপান্তর
একটি কল → তিনটি ক্যালেন্ডার। ইনপুট তারিখ একই সাথে তিনটি ক্যালেন্ডারে সমাধান করা হয়।

#### ② ত্রিগুণ লিপি আউটপুট
প্রতিটি তারিখ উপাদান তিনটি লিপিতে সাথে সাথে পাওয়া যায়। আপনাকে কখনও ম্যানুয়ালি অনুলিপি করতে হবে না।

#### ③ স্মার্ট বাংলা তারিখ প্রত্যয়
আনুষ্ঠানিক বাংলা লেখায় ব্যবহৃত ক্রমবাচক প্রত্যয় সঠিকভাবে প্রয়োগ করে। এই প্রত্যয় লজিক **তিনটি ক্যালেন্ডারেই** স্বাধীনভাবে প্রয়োগ হয়।

#### ④ ৬৩টি পাবলিক গ্লোবাল ভেরিয়েবল
ইঞ্জিন প্রতিটি ক্যালেন্ডার × প্রতিটি উপাদান × প্রতিটি লিপির সমন্বয় কভার করে **৬৩টি নামযুক্ত পাবলিক ভেরিয়েবল** লেখে।

#### ⑤ বাহ্যিক রিসোর্স ফাইল
সমস্ত লোকেল ডেটা `ABID_DateENGINE_resources.txt`-এ থাকে।

#### ⑥ লিপ ইয়ার সচেতনতা
ফাল্গুন মাসের ভেতরে পড়া গ্রেগরিয়ান লিপ ইয়ার বাউন্ডারি সঠিকভাবে পরিচালনা করে।

---

## 4. System Requirements | সিস্টেম প্রয়োজনীয়তা

### English

| Requirement | Minimum | Recommended |
|-------------|---------|-------------|
| Microsoft Word | 2007 | 2016 / 365 |
| Operating System | Windows 7 | Windows 10 / 11 |
| VBA Macros | Must be enabled | — |
| %APPDATA% write access | Required | — |

**Required COM components** (already present on standard Windows installs):
- `Scripting.FileSystemObject` — file existence check
- `ADODB.Stream` — UTF-8 file reading

> **💡 Tip:** To verify ADODB is available, open the VBA Editor (`Alt+F11`), go to **Tools → References**, and confirm *Microsoft ActiveX Data Objects 6.x Library* is listed (even unchecked, late binding still works as long as the DLL is registered on the system).

> **⚠️ Note:** This engine is Windows-only. It will not run on Mac Word because `ADODB.Stream` and `Environ("AppData")` are not available on macOS VBA.

### বাংলা

| প্রয়োজনীয়তা | ন্যূনতম | সুপারিশকৃত |
|--------------|---------|------------|
| Microsoft Word | ২০০৭ | ২০১৬ / ৩৬৫ |
| অপারেটিং সিস্টেম | Windows ৭ | Windows ১০ / ১১ |
| VBA ম্যাক্রো | সক্ষম থাকতে হবে | — |
| %APPDATA% লেখার অ্যাক্সেস | প্রয়োজন | — |

> **💡 টিপ:** ম্যাক্রো চালানোর আগে Word-এ Trust Center খুলুন (File → Options → Trust Center → Trust Center Settings → Macro Settings) এবং **"Enable all macros"** বা **"Disable all macros with notification"** বেছে নিন।

> **⚠️ সতর্কতা:** এই ইঞ্জিন শুধুমাত্র Windows-এ চলে। Mac Word-এ `ADODB.Stream` পাওয়া যায় না।

---

## 5. Installation | ইনস্টলেশন

### English

#### Step 1 — Enable Macros in Word
1. Open Word → **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Under **Macro Settings**, select *"Enable all macros"* (or at minimum, *"Disable with notification"*)
3. Click **OK**

#### Step 2 — Import the VBA Module
1. Open (or create) the Word document where you want to use the engine
2. Press **`Alt + F11`** to open the VBA Editor
3. In the Project Explorer, right-click your project → **Import File...**
4. Select `ABID__DateENGINE_v2.bas`
5. The module `ABID__DateENGINE` will appear under **Modules**

#### Step 3 — Create the Resource Directory
Open File Explorer or a Run dialog (`Win+R`) and create this exact path:

```
%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

Expanded on a typical system this becomes:
```
C:\Users\<YourUsername>\AppData\Roaming\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

You can create all nested folders at once with this PowerShell one-liner:
```powershell
New-Item -ItemType Directory -Force -Path "$env:APPDATA\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need"
```

#### Step 4 — Place the Resource File
Save `ABID_DateENGINE_resources.txt` (see [Section 6](#6-resource-file-structure--রিসোর্স-ফাইল-কাঠামো)) into the directory above.

> **⚠️ Critical:** The file **must** be saved as **UTF-8 encoding**. If you use Notepad, choose *"Save As"* and select **UTF-8** in the encoding dropdown. Using Notepad++ is strongly recommended.

#### Step 5 — Verify the Installation
Paste this quick-test macro into any module and run it:

```vba
Sub TestInstallation()
    Call ABID__DateENGINE(Format(Date, "DD/MM/YYYY"))
    MsgBox "Today in Bengali: " & Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
           Abid_DEngine_BanglaDate_monthName_bn & " " & _
           Abid_DEngine_BanglaDate_yearNum_bn & " " & _
           Abid_DEngine_BanglaDate_Era_bn, vbInformation, "Engine Test"
End Sub
```

If a message box shows a Bengali date — installation is complete! ✅

### বাংলা

#### ধাপ ১ — Word-এ ম্যাক্রো সক্ষম করুন
১. Word খুলুন → **File** → **Options** → **Trust Center** → **Trust Center Settings**
২. **Macro Settings**-এর অধীনে *"Enable all macros"* নির্বাচন করুন
৩. **OK** ক্লিক করুন

#### ধাপ ২ — VBA মডিউল ইম্পোর্ট করুন
১. **`Alt + F11`** চেপে VBA এডিটর খুলুন
২. প্রজেক্ট এক্সপ্লোরারে আপনার প্রজেক্টে রাইট-ক্লিক করুন → **Import File...**
৩. `ABID__DateENGINE_v2.bas` নির্বাচন করুন

#### ধাপ ৩ — রিসোর্স ডিরেক্টরি তৈরি করুন
Run ডায়ালগ (`Win+R`) খুলুন এবং এই পাথ দিয়ে ডিরেক্টরি তৈরি করুন:
```
%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

#### ধাপ ৪ — রিসোর্স ফাইল স্থাপন করুন
`ABID_DateENGINE_resources.txt` ফাইলটি উপরের ডিরেক্টরিতে **UTF-8 এনকোডিং** সহ সেভ করুন।

#### ধাপ ৫ — ইনস্টলেশন যাচাই করুন
উপরের ইংরেজি সেকশনের টেস্ট ম্যাক্রো চালান। বাংলা তারিখ দেখা গেলে ইনস্টলেশন সম্পন্ন! ✅

---

## 6. Resource File Structure | রিসোর্স ফাইল কাঠামো

### English

The resource file is a plain UTF-8 text file that stores all locale-specific strings. The engine reads it on every call to `ABID__DateENGINE`.

#### Format Rules

| Rule | Detail |
|------|--------|
| Encoding | UTF-8 (mandatory — Bengali characters will corrupt otherwise) |
| Section header | `[SECTION_NAME]` on its own line |
| Data rows | Three columns separated by pipe `\|` : `Bengali\|English\|Arabic` |
| Prefix row | Five values separated by commas: `লা,রা,ঠা,ই,শে` |
| Blank lines | Ignored |
| Line endings | CRLF or LF (both handled automatically) |

#### Complete File Template

```
[BANG_MONTHS]
বৈশাখ|Baishakh|بيشاخ
জ্যৈষ্ঠ|Jyaistha|جيشتها
আষাঢ়|Ashadh|اشاده
শ্রাবণ|Srabon|سرابون
ভাদ্র|Bhadro|بهادرو
আশ্বিন|Ashwin|اشوين
কার্তিক|Kartik|كارتيك
অগ্রহায়ণ|Ogrohayon|اغروهايون
পৌষ|Poush|پوش
মাঘ|Magh|ماغ
ফাল্গুন|Falgun|فالغون
চৈত্র|Choitro|چيترو

[GREGORIAN_MONTHS]
জানুয়ারি|January|يناير
ফেব্রুয়ারি|February|فبراير
মার্চ|March|مارس
এপ্রিল|April|ابريل
মে|May|مايو
জুন|June|يونيو
জুলাই|July|يوليو
আগস্ট|August|اغسطس
সেপ্টেম্বর|September|سبتمبر
অক্টোবর|October|اكتوبر
নভেম্বর|November|نوفمبر
ডিসেম্বর|December|ديسمبر

[HIJRI_MONTHS]
মহররম|Muharram|محرم
সফর|Safar|صفر
রবিউল আউয়াল|Rabi' al-awwal|ربيع الأول
রবিউস সানি|Rabi' al-thani|ربيع الثاني
জমাদিউল আউয়াল|Jumada al-awwal|جمادى الأولى
জমাদিউস সানি|Jumada al-thani|جمادى الثانية
রজব|Rajab|رجب
শা'বান|Sha'ban|شعبان
রমজান|Ramadan|رمضان
শাওয়াল|Shawwal|شوال
জিলক্বদ|Dhu al-Qi'dah|ذو القعدة
জিলহজ্জ|Dhu al-Hijjah|ذو الحجة

[ERAS]
বঙ্গাব্দ|Bangabda|بنغابدا
খ্রিস্টাব্দ|AD|ميلادي
হিজরি|AH|هجري

[WEEKDAYS]
রবিবার|Sunday|الأحد
সোমবার|Monday|الاثنين
মঙ্গলবার|Tuesday|الثلاثاء
বুধবার|Wednesday|الأربعاء
বৃহস্পতিবার|Thursday|الخميس
শুক্রবার|Friday|الجمعة
শনিবার|Saturday|السبت

[NUMERALS]
০|0|٠
১|1|١
২|2|٢
৩|3|٣
৪|4|٤
৫|5|٥
৬|6|٦
৭|7|٧
৮|8|٨
৯|9|٩

[BANG_DATE_NUM_PREFIX]
লা,রা,ঠা,ই,শে
```

> **💡 Tip:** The order of rows within each section is critical. `[BANG_MONTHS]` row 1 = Baishakh, row 2 = Jyaistha, etc. Never add or remove rows from a section, or the month index will shift and produce wrong dates.

> **💡 Tip:** The `[WEEKDAYS]` section starts with Sunday (index 1) because VBA's `Weekday()` function uses `vbSunday = 1` as the first day.

### বাংলা

রিসোর্স ফাইলটি একটি সাধারণ UTF-8 টেক্সট ফাইল যা সমস্ত লোকেল-নির্দিষ্ট স্ট্রিং সংরক্ষণ করে।

> **💡 টিপ:** প্রতিটি সেকশনের ভেতরের সারির ক্রম অত্যন্ত গুরুত্বপূর্ণ। `[BANG_MONTHS]`-এর ১ম সারি = বৈশাখ, ২য় সারি = জ্যৈষ্ঠ ইত্যাদি। কোনো সেকশন থেকে সারি যোগ বা বিয়োগ করবেন না, নয়তো মাসের ইন্ডেক্স পরিবর্তিত হয়ে ভুল তারিখ আসবে।

> **💡 টিপ:** `[WEEKDAYS]` সেকশন রবিবার দিয়ে শুরু হয় কারণ VBA-র `Weekday()` ফাংশন `vbSunday = 1` ব্যবহার করে।

---

## 7. Complete Variable Reference | সম্পূর্ণ ভেরিয়েবল রেফারেন্স

### English

All variables follow this naming pattern:

```
Abid_DEngine_[Calendar]_[Component]_[Script]
```

Where:
- **Calendar** = `BanglaDate` | `EnglishDate` | `ArabicDate`
- **Component** = `dayNum` | `monthNum` | `yearNum` | `weekName` | `monthName` | `Era`
- **Script** = `bn` (Bengali) | `en` (English) | `ar` (Arabic)
- **Special** = `dayNum_bn_pre` (Bengali numeral + ordinal suffix)

> **Example:** `Abid_DEngine_BanglaDate_monthName_ar` → the Bengali calendar's month name written in Arabic script → `بيشاخ`

---

### Group A — BanglaDate Variables (বাংলা ক্যালেন্ডার ভেরিয়েবল)

#### Day Numbers (তারিখ নম্বর)

| Variable | Script | Description | Example (15 Apr 2024 → 2 Baishakh 1431) |
|----------|--------|-------------|----------------------------------------|
| `Abid_DEngine_BanglaDate_dayNum_bn` | Bengali | Day number in Bengali numerals | `২` |
| `Abid_DEngine_BanglaDate_dayNum_bn_pre` | Bengali+suffix | Day number + ordinal suffix | `২রা` |
| `Abid_DEngine_BanglaDate_dayNum_en` | English | Day number in English numerals | `2` |
| `Abid_DEngine_BanglaDate_dayNum_ar` | Arabic | Day number in Arabic-Indic numerals | `٢` |

#### Month Numbers (মাস নম্বর)

| Variable | Script | Description | Example |
|----------|--------|-------------|---------|
| `Abid_DEngine_BanglaDate_monthNum_bn` | Bengali | Month number in Bengali numerals | `১` |
| `Abid_DEngine_BanglaDate_monthNum_en` | English | Month number in English numerals | `1` |
| `Abid_DEngine_BanglaDate_monthNum_ar` | Arabic | Month number in Arabic-Indic numerals | `١` |

#### Year Numbers (বছর নম্বর)

| Variable | Script | Description | Example |
|----------|--------|-------------|---------|
| `Abid_DEngine_BanglaDate_yearNum_bn` | Bengali | Year in Bengali numerals | `১৪৩১` |
| `Abid_DEngine_BanglaDate_yearNum_en` | English | Year in English numerals | `1431` |
| `Abid_DEngine_BanglaDate_yearNum_ar` | Arabic | Year in Arabic-Indic numerals | `١٤٣١` |

#### Week Names (সপ্তাহের নাম)

| Variable | Script | Example |
|----------|--------|---------|
| `Abid_DEngine_BanglaDate_weekName_bn` | Bengali | `সোমবার` |
| `Abid_DEngine_BanglaDate_weekName_en` | English | `Monday` |
| `Abid_DEngine_BanglaDate_weekName_ar` | Arabic | `الاثنين` |

#### Month Names (মাসের নাম)

| Variable | Script | Example |
|----------|--------|---------|
| `Abid_DEngine_BanglaDate_monthName_bn` | Bengali | `বৈশাখ` |
| `Abid_DEngine_BanglaDate_monthName_en` | English | `Baishakh` |
| `Abid_DEngine_BanglaDate_monthName_ar` | Arabic | `بيشاخ` |

#### Era Designations (যুগ পদবী)

| Variable | Script | Example |
|----------|--------|---------|
| `Abid_DEngine_BanglaDate_Era_bn` | Bengali | `বঙ্গাব্দ` |
| `Abid_DEngine_BanglaDate_Era_en` | English | `Bangabda` |
| `Abid_DEngine_BanglaDate_Era_ar` | Arabic | `بنغابدا` |

---

### Group B — EnglishDate Variables (গ্রেগরিয়ান ক্যালেন্ডার ভেরিয়েবল)

Same structure as BanglaDate. Example values for 15 April 2024:

| Variable | Example Value |
|----------|---------------|
| `Abid_DEngine_EnglishDate_dayNum_bn` | `১৫` |
| `Abid_DEngine_EnglishDate_dayNum_bn_pre` | `১৫ই` |
| `Abid_DEngine_EnglishDate_dayNum_en` | `15` |
| `Abid_DEngine_EnglishDate_dayNum_ar` | `١٥` |
| `Abid_DEngine_EnglishDate_monthNum_bn` | `৪` |
| `Abid_DEngine_EnglishDate_monthNum_en` | `4` |
| `Abid_DEngine_EnglishDate_monthNum_ar` | `٤` |
| `Abid_DEngine_EnglishDate_yearNum_bn` | `২০২৪` |
| `Abid_DEngine_EnglishDate_yearNum_en` | `2024` |
| `Abid_DEngine_EnglishDate_yearNum_ar` | `٢٠٢٤` |
| `Abid_DEngine_EnglishDate_weekName_bn` | `সোমবার` |
| `Abid_DEngine_EnglishDate_weekName_en` | `Monday` |
| `Abid_DEngine_EnglishDate_weekName_ar` | `الاثنين` |
| `Abid_DEngine_EnglishDate_monthName_bn` | `এপ্রিল` |
| `Abid_DEngine_EnglishDate_monthName_en` | `April` |
| `Abid_DEngine_EnglishDate_monthName_ar` | `ابريل` |
| `Abid_DEngine_EnglishDate_Era_bn` | `খ্রিস্টাব্দ` |
| `Abid_DEngine_EnglishDate_Era_en` | `AD` |
| `Abid_DEngine_EnglishDate_Era_ar` | `ميلادي` |

---

### Group C — ArabicDate Variables (হিজরি ক্যালেন্ডার ভেরিয়েবল)

Same structure. Example values for 15 April 2024 (= 6 Shawwal 1445 AH):

| Variable | Example Value |
|----------|---------------|
| `Abid_DEngine_ArabicDate_dayNum_bn` | `৬` |
| `Abid_DEngine_ArabicDate_dayNum_bn_pre` | `৬ই` |
| `Abid_DEngine_ArabicDate_dayNum_en` | `6` |
| `Abid_DEngine_ArabicDate_dayNum_ar` | `٦` |
| `Abid_DEngine_ArabicDate_monthNum_bn` | `১০` |
| `Abid_DEngine_ArabicDate_monthNum_en` | `10` |
| `Abid_DEngine_ArabicDate_monthNum_ar` | `١٠` |
| `Abid_DEngine_ArabicDate_yearNum_bn` | `১৪৪৫` |
| `Abid_DEngine_ArabicDate_yearNum_en` | `1445` |
| `Abid_DEngine_ArabicDate_yearNum_ar` | `١٤٤٥` |
| `Abid_DEngine_ArabicDate_weekName_bn` | `সোমবার` |
| `Abid_DEngine_ArabicDate_weekName_en` | `Monday` |
| `Abid_DEngine_ArabicDate_weekName_ar` | `الاثنين` |
| `Abid_DEngine_ArabicDate_monthName_bn` | `শাওয়াল` |
| `Abid_DEngine_ArabicDate_monthName_en` | `Shawwal` |
| `Abid_DEngine_ArabicDate_monthName_ar` | `شوال` |
| `Abid_DEngine_ArabicDate_Era_bn` | `হিজরি` |
| `Abid_DEngine_ArabicDate_Era_en` | `AH` |
| `Abid_DEngine_ArabicDate_Era_ar` | `هجري` |

> **💡 Note:** Week name variables (`weekName_*`) are identical across all three calendar groups for the same date. The engine deliberately populates all nine so your code does not need to know which calendar group to pull the weekday from.

### বাংলা — ভেরিয়েবল নামকরণ দ্রুত রেফারেন্স

```
Abid_DEngine_[ক্যালেন্ডার]_[উপাদান]_[লিপি]
             │                │          │
             ▼                ▼          ▼
  BanglaDate          dayNum          bn (বাংলা সংখ্যা)
  EnglishDate         monthNum        en (ইংরেজি সংখ্যা/লিখন)
  ArabicDate          yearNum         ar (আরবি-ইন্ডিক)
                      weekName
                      monthName
                      Era
                      dayNum_bn_pre  ← বিশেষ: প্রত্যয় সহ
```

---

## 8. Function Reference | ফাংশন রেফারেন্স

### English

#### Public Sub `ABID__DateENGINE`

```vba
Public Sub ABID__DateENGINE(ByVal InputDateStr As String)
```

**Purpose:** The single entry-point of the entire engine. Call this once with a date string; all 63 global variables are populated on return.

**Parameter:**
- `InputDateStr` — A date string in **`DD/MM/YYYY`** format. Leading zeros are required: `"05/04/2024"`, not `"5/4/2024"`.

**Returns:** Nothing (Side effect: writes to 63 public variables)

**Error conditions:**
- If the string cannot be split into 3 parts by `/` → shows `MsgBox` and exits early
- If the resource file is missing → shows `MsgBox` with the expected path and exits early

**Typical call:**
```vba
Call ABID__DateENGINE("15/04/2024")
' — or —
Call ABID__DateENGINE(Format(Date, "DD/MM/YYYY"))   ' today's date
```

---

#### Private Function `LoadDateResources`

```vba
Private Function LoadDateResources(
    ByRef arrBMonths()   As String,   ' Bengali month names   (12 × 3)
    ByRef arrGMonths()   As String,   ' Gregorian month names (12 × 3)
    ByRef arrHMonths()   As String,   ' Hijri month names     (12 × 3)
    ByRef arrEras()      As String,   ' Era strings           ( 3 × 3)
    ByRef arrWeeks()     As String,   ' Weekday names         ( 7 × 3)
    ByRef arrNums()      As String,   ' Digit map             (0–9 × 3)
    ByRef arrBPrefixes() As String    ' Ordinal suffixes      (1–5)
) As Boolean
```

**Purpose:** Opens the UTF-8 resource file with `ADODB.Stream`, parses all `[SECTION]` blocks, and fills the seven arrays passed in by reference.

**Returns:** `True` on success, `False` if the file is not found (also shows a `MsgBox`).

**Internal flow:**
1. Builds the file path from `Environ("AppData")`
2. Tries an optional `.txt` extension fallback
3. Reads the full file as UTF-8 text
4. Normalises line endings (`vbCrLf` → `vbLf`)
5. Iterates lines, switches `currentSection` on `[HEADER]` lines
6. Parses `|`-delimited rows into 2D arrays
7. Parses the comma-delimited prefix row into `arrBPrefixes(1..5)`

---

#### Private Function `ToLocalizedNum`

```vba
Private Function ToLocalizedNum(
    ByVal strNum  As String,
    ByVal langIdx As Integer,   ' 1=Bengali, 2=English, 3=Arabic
    ByRef arrNums() As String
) As String
```

**Purpose:** Converts a string of English digits (0–9) to the target script's digits by iterating through `arrNums` and doing character-by-character `Replace` calls.

**Example:**
```vba
ToLocalizedNum("2024", 1, arrNums)  ' → "২০২৪"
ToLocalizedNum("2024", 3, arrNums)  ' → "٢٠٢٤"
ToLocalizedNum("2024", 2, arrNums)  ' → "2024"  (no change)
```

### বাংলা — ফাংশন সংক্ষিপ্ত বিবরণ

| ফাংশন | প্রকার | উদ্দেশ্য |
|--------|--------|---------|
| `ABID__DateENGINE` | Public Sub | প্রধান এন্ট্রি পয়েন্ট — একটি তারিখ নিয়ে ৬৩টি ভেরিয়েবল পূর্ণ করে |
| `LoadDateResources` | Private Function | UTF-8 রিসোর্স ফাইল পড়ে ৭টি অ্যারে পূর্ণ করে |
| `ToLocalizedNum` | Private Function | ইংরেজি সংখ্যাকে লক্ষ্য লিপির সংখ্যায় রূপান্তর করে |

---

## 9. How the Algorithm Works | অ্যালগরিদম কিভাবে কাজ করে

### English

Understanding the internals will help you debug edge cases and extend the engine.

#### Step 1 — Parse Input
The input `"15/04/2024"` is split by `/` and assembled into a VBA `Date` object via `DateSerial(2024, 4, 15)`. This is the canonical date object used throughout.

#### Step 2 — Gregorian Extraction
Straightforward: `Day(dt)`, `Month(dt)`, `Year(dt)`. VBA's default calendar is Gregorian.

#### Step 3 — Hijri Conversion
VBA has a built-in calendar mode switch:
```vba
VBA.Calendar = vbCalHijri   ' switch to lunar calendar
hDay   = Day(dt)
hMonth = Month(dt)
hYear  = Year(dt)
VBA.Calendar = vbCalGreg    ' switch back immediately
```
The `dt` object holds the same point in time; the calendar switch just changes how VBA interprets `Day()/Month()/Year()`. The switch-back is **critical** — leaving it as Hijri would corrupt all subsequent Gregorian operations.

> **⚠️ Important:** VBA's built-in Hijri conversion follows the **tabular/calculated** Hijri calendar. It may differ by ±1 day from countries that use moon-sighting (رؤية الهلال). This is a known limitation of the VBA runtime.

#### Step 4 — Bengali Calendar Calculation (Most Complex)

The **official Bangladesh Bengali calendar** (implemented by the Bangla Academy and mandated by the government) uses a **solar, fixed calendar** that does *not* drift with the lunar Hijri calendar.

**Year calculation:**
```
If date is before April 14:  BengaliYear = GregorianYear - 594
If date is April 14 or later: BengaliYear = GregorianYear - 593
```

The -593/-594 offset comes from the Bangabda epoch (Bengali year 1 = 594 AD).

**Month and day calculation** — Fixed-rule mapping:

| Gregorian Range | Bengali Month | Day Offset Formula |
|-----------------|---------------|--------------------|
| Apr 14 – May 14 | বৈশাখ (1) | `bDay = gDay - 13` |
| May 15 – Jun 14 | জ্যৈষ্ঠ (2) | `bDay = gDay - 14` |
| Jun 15 – Jul 15 | আষাঢ় (3) | `bDay = gDay - 14` |
| Jul 16 – Aug 15 | শ্রাবণ (4) | `bDay = gDay - 15` |
| Aug 16 – Sep 15 | ভাদ্র (5) | `bDay = gDay - 15` |
| Sep 16 – Oct 16 | আশ্বিন (6) | `bDay = gDay - 15` |
| Oct 17 – Nov 15 | কার্তিক (7) | `bDay = gDay - 16` |
| Nov 16 – Dec 15 | অগ্রহায়ণ (8) | `bDay = gDay - 15` |
| Dec 16 – Jan 14 | পৌষ (9) | `bDay = gDay - 15` |
| Jan 15 – Feb 13 | মাঘ (10) | `bDay = gDay - 14` |
| Feb 14 – Mar 14 | ফাল্গুন (11) | `bDay = gDay - 13` (±leap) |
| Mar 15 – Apr 13 | চৈত্র (12) | `bDay = gDay - 14` |

> **Leap Year Edge Case:** In a Gregorian leap year, February has 29 days, which pushes the Falgun boundary one day forward. The engine checks `isLeapYear` and adds 1 to the offset for the March 1–14 span.

**Month lengths (reference):**

| Months | Length |
|--------|--------|
| বৈশাখ to ভাদ্র (1–5) | 31 days each |
| আশ্বিন to ফাল্গুন (6–11) | 30 days each |
| চৈত্র (12) | 30 days (31 in Bengali leap year) |

#### Step 5 — Numeral Localization
`ToLocalizedNum` does sequential `Replace()` calls on the English digit string, substituting each `0-9` character with its Bengali or Arabic-Indic counterpart from `arrNums`. Because it replaces single characters, multi-digit numbers are handled correctly.

#### Step 6 — Suffix Application
A `Select Case` block maps the numeric day to one of five suffixes from `arrBPrefixes`. The suffixed form is stored in the `_bn_pre` variable while the raw number stays in `_bn`, giving callers the choice.

### বাংলা — অ্যালগরিদম ব্যাখ্যা

#### বাংলা ক্যালেন্ডার গণনা (সবচেয়ে জটিল)

বাংলা একাডেমি প্রণীত এবং বাংলাদেশ সরকার অনুমোদিত **অফিশিয়াল বাংলা ক্যালেন্ডার** একটি **সৌর, নির্দিষ্ট ক্যালেন্ডার** যা চন্দ্র হিজরি ক্যালেন্ডারের মতো ভাসমান নয়।

**বছর গণনা সূত্র:**
```
১৪ এপ্রিলের আগে:    বাংলা বছর = গ্রেগরিয়ান বছর − ৫৯৪
১৪ এপ্রিল বা পরে:   বাংলা বছর = গ্রেগরিয়ান বছর − ৫৯৩
```

এই −৫৯৩/−৫৯৪ অফসেট বঙ্গাব্দের উৎস বিন্দু থেকে আসে (বঙ্গাব্দ ১ = ৫৯৪ খ্রিস্টাব্দ)।

> **⚠️ গুরুত্বপূর্ণ:** VBA-এর বিল্ট-ইন হিজরি রূপান্তর **গণনামূলক/সারণী** হিজরি ক্যালেন্ডার অনুসরণ করে। চাঁদ দেখার (রুয়াতুল হিলাল) উপর নির্ভরশীল দেশগুলির সাথে ±১ দিনের পার্থক্য হতে পারে।

---

## 10. Usage Examples & Recipes | ব্যবহারের উদাহরণ ও রেসিপি

### English

---

#### Recipe 1 — Today's Date in Bengali (সবচেয়ে সহজ)

```vba
Sub Recipe_TodayInBengali()
    Call ABID__DateENGINE(Format(Date, "DD/MM/YYYY"))
    
    MsgBox Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
           Abid_DEngine_BanglaDate_monthName_bn & " " & _
           Abid_DEngine_BanglaDate_yearNum_bn & " " & _
           Abid_DEngine_BanglaDate_Era_bn
    ' Output example: ২রা বৈশাখ ১৪৩১ বঙ্গাব্দ
End Sub
```

---

#### Recipe 2 — Three-Calendar Letterhead (তিন-ক্যালেন্ডার লেটারহেড)

A classic use case: government and official letters in Bangladesh often require all three calendar dates.

```vba
Sub Recipe_TripleCalendarLetterhead()
    ' -------------------------------------------------------
    ' Insert a formatted three-calendar date block at cursor
    ' -------------------------------------------------------
    Call ABID__DateENGINE(Format(Date, "DD/MM/YYYY"))
    
    Dim banglaLine   As String
    Dim englishLine  As String
    Dim hijriLine    As String
    Dim weekLine     As String
    
    ' Bengali line: ২রা বৈশাখ ১৪৩১ বঙ্গাব্দ
    banglaLine = Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                 Abid_DEngine_BanglaDate_monthName_bn & " " & _
                 Abid_DEngine_BanglaDate_yearNum_bn & " " & _
                 Abid_DEngine_BanglaDate_Era_bn
    
    ' English line: 15 April 2024 AD
    englishLine = Abid_DEngine_EnglishDate_dayNum_en & " " & _
                  Abid_DEngine_EnglishDate_monthName_en & " " & _
                  Abid_DEngine_EnglishDate_yearNum_en & " " & _
                  Abid_DEngine_EnglishDate_Era_en
    
    ' Hijri line: ৬ই শাওয়াল ১৪৪৫ হিজরি
    hijriLine  = Abid_DEngine_ArabicDate_dayNum_bn_pre & " " & _
                 Abid_DEngine_ArabicDate_monthName_bn & " " & _
                 Abid_DEngine_ArabicDate_yearNum_bn & " " & _
                 Abid_DEngine_ArabicDate_Era_bn
    
    ' Weekday
    weekLine = "বার: " & Abid_DEngine_BanglaDate_weekName_bn & _
               " / " & Abid_DEngine_BanglaDate_weekName_en
    
    ' Insert into document
    With Selection
        .TypeText "তারিখ:" & vbCrLf
        .TypeText "বাংলা  — " & banglaLine  & vbCrLf
        .TypeText "English — " & englishLine & vbCrLf
        .TypeText "হিজরি  — " & hijriLine   & vbCrLf
        .TypeText weekLine & vbCrLf
    End With
End Sub
```

---

#### Recipe 3 — Populate a Word Table (টেবিল পূর্ণ করুন)

```vba
Sub Recipe_FillDateTable()
    ' Assumes a 3-row × 2-column table is selected / at cursor
    Call ABID__DateENGINE("01/01/2025")
    
    Dim tbl As Table
    Set tbl = ActiveDocument.Tables(1)
    
    ' Header labels
    tbl.Cell(1, 1).Range.Text = "ক্যালেন্ডার"
    tbl.Cell(2, 1).Range.Text = "বাংলা তারিখ"
    tbl.Cell(3, 1).Range.Text = "English Date"
    tbl.Cell(4, 1).Range.Text = "التاريخ الهجري"
    
    ' Date values
    tbl.Cell(2, 2).Range.Text = Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                                 Abid_DEngine_BanglaDate_monthName_bn & " " & _
                                 Abid_DEngine_BanglaDate_yearNum_bn
    tbl.Cell(3, 2).Range.Text = Abid_DEngine_EnglishDate_monthName_en & " " & _
                                 Abid_DEngine_EnglishDate_dayNum_en & ", " & _
                                 Abid_DEngine_EnglishDate_yearNum_en
    tbl.Cell(4, 2).Range.Text = Abid_DEngine_ArabicDate_dayNum_ar & " " & _
                                 Abid_DEngine_ArabicDate_monthName_ar & " " & _
                                 Abid_DEngine_ArabicDate_yearNum_ar
End Sub
```

---

#### Recipe 4 — Dynamic Date from a Bookmark (বুকমার্ক থেকে তারিখ)

```vba
Sub Recipe_DateFromBookmark()
    ' Reads a date stored in a document bookmark named "InputDate"
    Dim bm As Bookmark
    Set bm = ActiveDocument.Bookmarks("InputDate")
    Dim inputStr As String
    inputStr = bm.Range.Text   ' expects "DD/MM/YYYY"
    
    If inputStr = "" Then
        MsgBox "Bookmark 'InputDate' is empty!", vbExclamation
        Exit Sub
    End If
    
    Call ABID__DateENGINE(inputStr)
    
    ' Now write the Bengali result to another bookmark
    ActiveDocument.Bookmarks("BengaliDateOutput").Range.Text = _
        Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
        Abid_DEngine_BanglaDate_monthName_bn & " " & _
        Abid_DEngine_BanglaDate_yearNum_bn & " " & _
        Abid_DEngine_BanglaDate_Era_bn
End Sub
```

---

#### Recipe 5 — Batch Date Conversion Loop (একাধিক তারিখ রূপান্তর)

```vba
Sub Recipe_BatchConversion()
    ' Convert a list of dates and print results to Immediate Window
    Dim dates() As String
    dates = Split("01/01/2024,14/04/2024,01/06/2024,31/12/2024", ",")
    
    Dim i As Integer
    For i = 0 To UBound(dates)
        Call ABID__DateENGINE(dates(i))
        Debug.Print dates(i) & " → " & _
            Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
            Abid_DEngine_BanglaDate_monthName_bn & " " & _
            Abid_DEngine_BanglaDate_yearNum_bn & " " & _
            Abid_DEngine_BanglaDate_Era_bn
    Next i
    ' Press Ctrl+G in VBA Editor to see output
End Sub
```

> **💡 Performance Note:** Each loop iteration re-reads the resource file. For large batches, see [Section 12](#12-performance-considerations--পারফরম্যান্স-বিবেচনা) for a caching pattern.

---

#### Recipe 6 — Full Arabic-Script Hijri Date

```vba
Sub Recipe_FullArabicHijriDate()
    Call ABID__DateENGINE("15/04/2024")
    
    ' Arabic format: day-month-year all in Arabic script
    Dim arabicDate As String
    arabicDate = Abid_DEngine_ArabicDate_dayNum_ar & " " & _
                 Abid_DEngine_ArabicDate_monthName_ar & " " & _
                 Abid_DEngine_ArabicDate_yearNum_ar & " " & _
                 Abid_DEngine_ArabicDate_Era_ar
    
    Debug.Print arabicDate
    ' Output: ٦ شوال ١٤٤٥ هجري
End Sub
```

---

#### Recipe 7 — Reusable Formatter Function (পুনর্ব্যবহারযোগ্য ফরম্যাটার)

Create a helper function that accepts parameters so you don't repeat yourself:

```vba
' -------------------------------------------------------
' FormatDate: Flexible date formatter using the engine.
' calendar : "Bangla" | "English" | "Arabic"
' style    : "short" (day month year) | "long" (weekday, day month year era)
' script   : "bn" | "en" | "ar"
' -------------------------------------------------------
Function FormatDate(dateStr As String, calendar As String, _
                    style As String, script As String) As String
    Call ABID__DateENGINE(dateStr)
    
    Dim d As String, m As String, y As String, w As String, e As String
    
    Select Case calendar
        Case "Bangla"
            Select Case script
                Case "bn": d = Abid_DEngine_BanglaDate_dayNum_bn_pre
                           m = Abid_DEngine_BanglaDate_monthName_bn
                           y = Abid_DEngine_BanglaDate_yearNum_bn
                           w = Abid_DEngine_BanglaDate_weekName_bn
                           e = Abid_DEngine_BanglaDate_Era_bn
                Case "en": d = Abid_DEngine_BanglaDate_dayNum_en
                           m = Abid_DEngine_BanglaDate_monthName_en
                           y = Abid_DEngine_BanglaDate_yearNum_en
                           w = Abid_DEngine_BanglaDate_weekName_en
                           e = Abid_DEngine_BanglaDate_Era_en
            End Select
        Case "English"
            Select Case script
                Case "bn": d = Abid_DEngine_EnglishDate_dayNum_bn_pre
                           m = Abid_DEngine_EnglishDate_monthName_bn
                           y = Abid_DEngine_EnglishDate_yearNum_bn
                           w = Abid_DEngine_EnglishDate_weekName_bn
                           e = Abid_DEngine_EnglishDate_Era_bn
                Case "en": d = Abid_DEngine_EnglishDate_dayNum_en
                           m = Abid_DEngine_EnglishDate_monthName_en
                           y = Abid_DEngine_EnglishDate_yearNum_en
                           w = Abid_DEngine_EnglishDate_weekName_en
                           e = Abid_DEngine_EnglishDate_Era_en
            End Select
        Case "Arabic"
            Select Case script
                Case "ar": d = Abid_DEngine_ArabicDate_dayNum_ar
                           m = Abid_DEngine_ArabicDate_monthName_ar
                           y = Abid_DEngine_ArabicDate_yearNum_ar
                           w = Abid_DEngine_ArabicDate_weekName_ar
                           e = Abid_DEngine_ArabicDate_Era_ar
                Case "bn": d = Abid_DEngine_ArabicDate_dayNum_bn_pre
                           m = Abid_DEngine_ArabicDate_monthName_bn
                           y = Abid_DEngine_ArabicDate_yearNum_bn
                           w = Abid_DEngine_ArabicDate_weekName_bn
                           e = Abid_DEngine_ArabicDate_Era_bn
            End Select
    End Select
    
    If style = "long" Then
        FormatDate = w & ", " & d & " " & m & " " & y & " " & e
    Else
        FormatDate = d & " " & m & " " & y
    End If
End Function

' --- Usage ---
' MsgBox FormatDate("15/04/2024", "Bangla",  "long", "bn")
' → সোমবার, ২রা বৈশাখ ১৪৩১ বঙ্গাব্দ
'
' MsgBox FormatDate("15/04/2024", "English", "short", "en")
' → 15 April 2024
'
' MsgBox FormatDate("15/04/2024", "Arabic",  "long", "ar")
' → الاثنين, ٦ شوال ١٤٤٥ هجري
```

---

#### Recipe 8 — Input Validation Wrapper (নিরাপদ কল)

```vba
Function SafeCallEngine(dateStr As String) As Boolean
    ' Returns True if call succeeded, False otherwise
    On Error GoTo Fail
    
    ' Validate format
    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) <> 2 Then GoTo Fail
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then GoTo Fail
    If CInt(parts(0)) < 1 Or CInt(parts(0)) > 31 Then GoTo Fail
    If CInt(parts(1)) < 1 Or CInt(parts(1)) > 12 Then GoTo Fail
    If Len(parts(2)) <> 4 Then GoTo Fail
    
    Call ABID__DateENGINE(dateStr)
    SafeCallEngine = True
    Exit Function
    
Fail:
    MsgBox "Invalid date: '" & dateStr & "'. Expected DD/MM/YYYY format.", vbExclamation
    SafeCallEngine = False
End Function

' Usage:
' If SafeCallEngine("15/04/2024") Then
'     Debug.Print Abid_DEngine_BanglaDate_dayNum_bn_pre
' End If
```

---

### বাংলা — উদাহরণ সংক্ষিপ্ত বিবরণ

| রেসিপি | উদ্দেশ্য |
|--------|---------|
| ১ | আজকের তারিখ বাংলায় দেখান |
| ২ | তিন-ক্যালেন্ডার লেটারহেড তৈরি করুন |
| ৩ | Word টেবিলে তারিখ পূর্ণ করুন |
| ৪ | বুকমার্ক থেকে তারিখ পড়ুন ও লিখুন |
| ৫ | একাধিক তারিখ লুপে রূপান্তর করুন |
| ৬ | সম্পূর্ণ আরবি হিজরি তারিখ তৈরি করুন |
| ৭ | পুনর্ব্যবহারযোগ্য ফরম্যাটার ফাংশন |
| ৮ | ইনপুট যাচাই করে নিরাপদভাবে ইঞ্জিন কল করুন |

---

## 11. Top Tips & Best Practices | শীর্ষ টিপস ও সেরা অনুশীলন

### English

#### 🏆 Tip 1 — Call Once, Read Many Times
Each call to `ABID__DateENGINE` reads the resource file from disk. If you need the same date in multiple places, call the engine **once**, save the results to local variables, then use those local variables everywhere.

```vba
' ✅ GOOD — one disk read
Call ABID__DateENGINE("15/04/2024")
Dim banglaDay   As String: banglaDay  = Abid_DEngine_BanglaDate_dayNum_bn_pre
Dim banglaMonth As String: banglaMonth = Abid_DEngine_BanglaDate_monthName_bn
Dim banglaYear  As String: banglaYear  = Abid_DEngine_BanglaDate_yearNum_bn

' Use banglaDay / banglaMonth / banglaYear repeatedly

' ❌ AVOID — multiple disk reads for same date
For i = 1 To 10
    Call ABID__DateENGINE("15/04/2024")  ' reads file 10 times!
    ' ...
Next i
```

---

#### 🏆 Tip 2 — Always Use `DD/MM/YYYY` with Leading Zeros
The engine splits by `/` and casts with `CInt`. `"5/4/2024"` works, but `"5/4/24"` will produce a wrong year. Always use 4-digit years and 2-digit day/month for clarity.

```vba
' ✅  Correct
Call ABID__DateENGINE("05/04/2024")

' ⚠️  Works but risky
Call ABID__DateENGINE("5/4/2024")

' ❌  Wrong year — will calculate year 24 instead of 2024
Call ABID__DateENGINE("05/04/24")
```

---

#### 🏆 Tip 3 — Format Today's Date Safely
Always use `Format(Date, "DD/MM/YYYY")` rather than concatenating `Day(Date)`, `Month(Date)` manually, because `Format` guarantees the leading zeros.

```vba
' ✅ Safe — always produces "05/04/2024" format
Call ABID__DateENGINE(Format(Date, "DD/MM/YYYY"))

' ⚠️ Risky — Day(Date) returns "5" not "05" for single-digit days
Call ABID__DateENGINE(Day(Date) & "/" & Month(Date) & "/" & Year(Date))
```

---

#### 🏆 Tip 4 — Use `_bn_pre` for Formal Bengali Text
For any formal Bengali document, always prefer `dayNum_bn_pre` (e.g., `২রা`) over raw `dayNum_bn` (e.g., `২`). The suffix is what distinguishes formal written Bengali from numeric-only output.

---

#### 🏆 Tip 5 — Wrap in Error Handler for Production
In production macros, always wrap the engine call in an error handler so that a missing resource file or a corrupt date doesn't crash the whole document workflow.

```vba
Sub ProductionSafeCall(inputDate As String)
    On Error GoTo ErrHandler
    Call ABID__DateENGINE(inputDate)
    On Error GoTo 0   ' reset error handling
    
    ' Continue with your logic here
    Exit Sub

ErrHandler:
    MsgBox "DateEngine Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "ABID DateEngine"
    On Error GoTo 0
End Sub
```

---

#### 🏆 Tip 6 — Resource File is Your Friend for Customisation
If you need to change a month name spelling (e.g., `Jyaistha` → `Jyeshtha`) or add a variant transliteration, just edit the resource file. No VBA code changes needed. Always keep a backup of the original resource file before editing.

---

#### 🏆 Tip 7 — Use Immediate Window for Quick Debugging
In the VBA Editor, press `Ctrl+G` to open the Immediate Window, then run:

```vba
Call ABID__DateENGINE("14/04/2024") : ? Abid_DEngine_BanglaDate_dayNum_bn_pre
```

This lets you spot-check any date without creating a full test macro.

---

#### 🏆 Tip 8 — Bengali Numerals vs Arabic-Indic: Know the Difference
The engine produces three numeral styles. Make sure you're using the right one for your context:

| Variable suffix | Output for "15" | Used for |
|-----------------|-----------------|---------|
| `_bn` | `১৫` | Bengali documents |
| `_en` | `15` | English documents |
| `_ar` | `١٥` | Arabic/Hijri contexts |

---

### বাংলা — শীর্ষ টিপস

| টিপ | সংক্ষিপ্ত বিবরণ |
|-----|----------------|
| ১ | একবার কল করুন, অনেকবার পড়ুন — প্রতিটি কল ডিস্ক থেকে ফাইল পড়ে |
| ২ | সবসময় `DD/MM/YYYY` ফরম্যাটে শূন্য-সহ লিখুন |
| ৩ | আজকের তারিখের জন্য `Format(Date, "DD/MM/YYYY")` ব্যবহার করুন |
| ৪ | আনুষ্ঠানিক বাংলার জন্য `_bn_pre` (প্রত্যয় সহ) ব্যবহার করুন |
| ৫ | প্রোডাকশনে সবসময় Error Handler যোগ করুন |
| ৬ | মাসের নাম পরিবর্তন করতে শুধু রিসোর্স ফাইল এডিট করুন, কোড নয় |
| ৭ | দ্রুত ডিবাগিংয়ের জন্য Immediate Window ব্যবহার করুন (`Ctrl+G`) |
| ৮ | `_bn`, `_en`, `_ar` — কোনটি কোন প্রসঙ্গে ব্যবহার করবেন তা জানুন |

---

## 12. Performance Considerations | পারফরম্যান্স বিবেচনা

### English

The main performance cost of the engine is the **file I/O** inside `LoadDateResources`. On every call to `ABID__DateENGINE`, the resource file is opened, fully read, parsed, and then discarded. For single-date use cases (one letterhead, one certificate) this is completely acceptable — the file is small and the operation takes milliseconds.

However, in **batch processing** scenarios, this can accumulate. Here is a caching pattern for heavy workloads:

#### Caching Pattern for Batch Operations

Expose the resource arrays as module-level variables so they are loaded only once:

```vba
' -------------------------------------------------------
' Module: DateEngineBatch
' Add this to a new module alongside the engine module.
' -------------------------------------------------------
Private m_BMonths()   As String
Private m_GMonths()   As String
Private m_HMonths()   As String
Private m_Eras()      As String
Private m_Weeks()     As String
Private m_Nums()      As String
Private m_BPrefixes() As String
Private m_Loaded      As Boolean

Sub PreloadResources()
    If Not m_Loaded Then
        m_Loaded = LoadDateResources(m_BMonths, m_GMonths, m_HMonths, _
                                     m_Eras, m_Weeks, m_Nums, m_BPrefixes)
    End If
End Sub
```

> **Note:** `LoadDateResources` is `Private` in the current engine. To implement this pattern properly, you would need to change it to `Friend` or refactor it — see [Section 14](#14-developer-notes--upgrade-roadmap--ডেভেলপার-নোট-ও-আপগ্রেড-রোডম্যাপ) for guidance.

#### Other Performance Tips

- Avoid calling `ABID__DateENGINE` inside `Document_Open` if the document loads frequently and the date rarely changes — instead call it lazily on first use.
- For very large batch jobs (thousands of dates), consider exporting the logic to a standalone VBScript or Python script instead of Word VBA.
- The `ToLocalizedNum` function loops 10 times per number string. This is negligible for normal use.

### বাংলা — পারফরম্যান্স

ইঞ্জিনের প্রধান পারফরম্যান্স খরচ হলো `LoadDateResources`-এর ভেতরের **ফাইল I/O**। প্রতিটি কলে রিসোর্স ফাইল খোলা, পড়া, পার্স করা এবং তারপর বাদ দেওয়া হয়। একক-তারিখের ব্যবহারের জন্য এটি সম্পূর্ণ গ্রহণযোগ্য — ফাইলটি ছোট এবং অপারেশনটি মিলিসেকেন্ড নেয়।

বড় ব্যাচ অপারেশনের জন্য উপরের ক্যাশিং প্যাটার্ন ব্যবহার করুন।

---

## 13. Troubleshooting Guide | সমস্যা সমাধান গাইড

### English

#### 🔴 Error: "Resource file not found at: ..."

**Cause:** The `.txt` resource file is missing or in the wrong folder.

**Fix:**
1. Open File Explorer and navigate to `%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\`
2. Confirm `ABID_DateENGINE_resources.txt` exists there
3. Check for typos in the filename (it is case-sensitive on some systems)
4. If the folder doesn't exist, see [Installation Step 3](#step-3--create-the-resource-directory)

---

#### 🔴 Error: Bengali / Arabic characters show as boxes or `????`

**Cause:** Resource file is not saved as UTF-8.

**Fix:**
1. Open the file in **Notepad++**
2. Menu: **Encoding → Convert to UTF-8** (not UTF-8 BOM)
3. Save (`Ctrl+S`)
4. Restart Word and try again

Alternatively in Windows Notepad (Win 10/11): *File → Save As → Encoding: UTF-8*

---

#### 🔴 Error: Wrong Bengali date calculated

**Cause A:** Input date is not in `DD/MM/YYYY` format (e.g., using `MM/DD/YYYY`).
**Fix:** Ensure day comes first: `"15/04/2024"` not `"04/15/2024"`.

**Cause B:** Single-digit day or month without leading zero and a short year.
**Fix:** Always use 4-digit year: `"05/04/2024"` not `"5/4/24"`.

---

#### 🔴 Error: All variables are empty after running

**Cause:** The engine exited early due to a silent error (resource parse failure, or date parse error that was dismissed).

**Fix:**
1. Open VBA Editor → **View → Immediate Window** (`Ctrl+G`)
2. Type `Call ABID__DateENGINE("15/04/2024")` and press Enter
3. Then type `? Abid_DEngine_BanglaDate_dayNum_bn` and press Enter
4. If still empty, check whether a MsgBox appeared and was dismissed

---

#### 🔴 Error: "ADODB.Stream" related runtime error

**Cause:** ActiveX Data Objects library DLL is not registered on the machine (unusual but possible on stripped-down Windows).

**Fix:**
1. In VBA Editor: **Tools → References**
2. Check whether *Microsoft ActiveX Data Objects 6.x Library* appears
3. If missing, search for `msado15.dll` on your system and register it:
   ```
   regsvr32 "C:\Program Files\Common Files\System\ado\msado15.dll"
   ```
   (Run Command Prompt as Administrator)

---

#### 🔴 Error: Hijri date is off by 1 day

**Cause:** This is expected behaviour. VBA's built-in Hijri calendar uses the **tabular/calculated** method. Countries that rely on moon-sighting may observe a different day.

**Not a bug** — it is a known limitation of `vbCalHijri`. See [Developer Notes](#14-developer-notes--upgrade-roadmap--ডেভেলপার-নোট-ও-আপগ্রেড-রোডম্যাপ) for upgrade possibilities.

---

#### 🔴 Error: Resource file sections not found / arrays all empty

**Cause:** Resource file has Windows CRLF line endings but they weren't normalised, OR the file has a BOM prefix that confuses section header detection.

**Fix:** Re-save the file with UTF-8 (no BOM) encoding using Notepad++:
- **Encoding → Encode in UTF-8** (not "UTF-8 BOM")

---

### বাংলা — সমস্যা দ্রুত রেফারেন্স

| সমস্যা | সম্ভাব্য কারণ | সমাধান |
|--------|--------------|--------|
| "Resource file not found" | ফাইল নেই বা ভুল পাথে | ডিরেক্টরি ও ফাইলের নাম যাচাই করুন |
| বাংলা অক্ষর বক্স দেখাচ্ছে | UTF-8 এনকোডিং নেই | Notepad++-এ UTF-8 এ কনভার্ট করুন |
| ভুল বাংলা তারিখ | তারিখ ফরম্যাট ভুল | `DD/MM/YYYY` নিশ্চিত করুন |
| ভেরিয়েবল খালি | ইঞ্জিন আগেই বের হয়েছে | VBA Immediate Window-এ ডিবাগ করুন |
| ADODB ত্রুটি | DLL রেজিস্টার্ড নয় | `msado15.dll` রেজিস্টার করুন |
| হিজরি তারিখ ১ দিন পার্থক্য | VBA-এর গণনা পদ্ধতি | এটি পরিচিত সীমাবদ্ধতা |

---

## 14. Developer Notes & Upgrade Roadmap | ডেভেলপার নোট ও আপগ্রেড রোডম্যাপ

### English

This section is intended for developers who want to extend, port, or significantly improve the engine. The codebase is deliberately minimal and readable — most upgrades are straightforward once you understand the architecture.

---

### 🛠 Upgrade Idea 1 — Resource Caching (High Priority)

**Problem:** The resource file is read from disk on every call. In batch scenarios this is wasteful.

**Solution:** Promote `LoadDateResources` from `Private` to `Friend` (or add a public `InitEngine` sub that loads once and stores into module-level arrays). Add a `m_Loaded As Boolean` flag.

```vba
' Proposed API
Public Sub InitEngine()
    ' Call this ONCE at startup
    If Not m_Loaded Then
        m_Loaded = LoadDateResources(m_BM, m_GM, m_HM, m_Er, m_Wk, m_Nu, m_Pre)
    End If
End Sub

Public Sub ABID__DateENGINE_Cached(ByVal InputDateStr As String)
    If Not m_Loaded Then Call InitEngine
    ' ... rest of engine using module-level arrays instead of local
End Sub
```

---

### 🛠 Upgrade Idea 2 — Add a Fourth Language (e.g., Devanagari / Urdu)

The resource file format is already designed for extension. Each data row is `Bengali|English|Arabic`. To add Hindi/Devanagari:

1. Add a 4th column: `বৈশাখ|Baishakh|بيشاخ|बैशाख`
2. Change `ReDim arrBMonths(1 To 12, 1 To 4)` (and same for other arrays)
3. Add new `Public` variables with `_hi` suffix
4. Update `ToLocalizedNum` caller to support `langIdx = 4`

The engine will then expose `Abid_DEngine_BanglaDate_monthName_hi` etc.

---

### 🛠 Upgrade Idea 3 — Add More Calendar Systems

**Persian (Solar Hijri / Shamsi):**
There is no built-in VBA support for the Persian calendar. You would need to implement the Persian calendar algorithm directly in VBA (the Jalali conversion formula is well-documented). Hook it in at Step 4 of the engine alongside the existing Bangla calculation.

**Julian Calendar:**
For historical date research, add a Julian calendar converter. The offset between Julian and Gregorian is predictable by century.

**Hebrew Calendar:**
VBA has no `vbCalHebrew` equivalent to `vbCalHijri`. A pure-VBA Hebrew calendar algorithm would need to be implemented.

---

### 🛠 Upgrade Idea 4 — Input Format Flexibility

Currently only `DD/MM/YYYY` is accepted. Add an overloaded version (or an optional `formatStr` parameter) that accepts:
- `YYYY-MM-DD` (ISO 8601)
- `MM/DD/YYYY` (US format)
- A VBA `Date` object directly (most natural for internal callers)

```vba
' Proposed signature
Public Sub ABID__DateENGINE_v3(ByVal InputDate As Variant, _
                                Optional ByVal fmt As String = "DD/MM/YYYY")
```

---

### 🛠 Upgrade Idea 5 — Return a Struct Instead of Global Variables

Global variables work fine within a single Word document, but they create tight coupling when the engine is included in a Word Template (`.dotm`) used by multiple documents. A cleaner pattern is to return a **User-Defined Type**:

```vba
Type DateResult
    BanglaDay_bn    As String
    BanglaDay_pre   As String
    BanglaMonth_bn  As String
    BanglaYear_bn   As String
    ' ... all 63 fields
End Type

Function GetDateResult(dateStr As String) As DateResult
    Call ABID__DateENGINE(dateStr)
    Dim r As DateResult
    r.BanglaDay_bn  = Abid_DEngine_BanglaDate_dayNum_bn
    r.BanglaDay_pre = Abid_DEngine_BanglaDate_dayNum_bn_pre
    ' ... fill struct
    GetDateResult = r
End Function
```

This makes the engine thread-safe (no shared global state) and easier to unit test.

---

### 🛠 Upgrade Idea 6 — Port to Excel / Access / Outlook VBA

The engine is in a `.bas` module, so it can be imported into **any VBA host**, not just Word. For Excel:
- Change `Selection.TypeText` references in examples to `Range` operations
- The engine core itself requires **no changes** — only caller code changes

For Outlook, the engine could auto-insert multi-calendar dates into email signatures.

---

### 🛠 Upgrade Idea 7 — Moon-Sighting Hijri (Advanced)

To fix the ±1 day Hijri discrepancy, replace the `vbCalHijri` block with an API call to a live Hijri date service:

```vba
' Concept (requires internet access from VBA via MSXML2)
Function GetHijriFromAPI(gregDate As String) As String
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", "https://api.aladhan.com/v1/gToH?date=" & gregDate, False
    xhr.Send
    ' Parse JSON response ...
    ' Return "DD/MM/YYYY" Hijri string
End Function
```

The Al-Adhan API (`api.aladhan.com`) is free and returns multiple Hijri calendar methods.

---

### 🛠 Upgrade Idea 8 — GUI UserForm

Add a Word UserForm that lets the user:
1. Pick a date from a calendar control
2. Preview all three calendar outputs in real time
3. Choose a format template and insert into the document

This would make the engine accessible to non-developer users who just need to insert dates without writing macros.

---

### 🛠 Upgrade Idea 9 — Unit Tests

Create a dedicated test module:

```vba
Sub RunAllTests()
    ' Test key boundary dates
    TestDate "14/04/2024", "bMonth", "1",    "1st of Baishakh should be month 1"
    TestDate "13/04/2024", "bMonth", "12",   "Last day before New Year = Choitro"
    TestDate "29/02/2024", "bDay",   "17",   "Leap day Feb 29 in Falgun"
    TestDate "01/01/2024", "gYear",  "2024", "Year passthrough"
    Debug.Print "All tests passed!"
End Sub

Sub TestDate(dateStr As String, varName As String, expected As String, msg As String)
    Call ABID__DateENGINE(dateStr)
    Dim actual As String
    Select Case varName
        Case "bMonth": actual = Abid_DEngine_BanglaDate_monthNum_en
        Case "bDay":   actual = Abid_DEngine_BanglaDate_dayNum_en
        Case "gYear":  actual = Abid_DEngine_EnglishDate_yearNum_en
    End Select
    If actual <> expected Then
        MsgBox "FAIL: " & msg & vbCrLf & "Expected: " & expected & ", Got: " & actual, vbCritical
    End If
End Sub
```

---

### 🛠 Upgrade Idea 10 — Bengali Leap Year (Bangabda)

The current engine uses the Gregorian leap year to adjust the Falgun boundary. A more complete implementation would also correctly determine whether the **Bengali year itself** is a leap year (affecting Choitro having 31 days vs 30). The rule for Bengali leap years follows a modified version of the Gregorian leap year rule applied to the Bengali year number.

---

### বাংলা — ডেভেলপার নোট সংক্ষিপ্ত তালিকা

| আইডিয়া | অগ্রাধিকার | কঠিনতা |
|---------|-----------|---------|
| রিসোর্স ক্যাশিং | উচ্চ | সহজ |
| চতুর্থ ভাষা (দেবনাগরী/উর্দু) | মাঝারি | সহজ |
| আরও ক্যালেন্ডার (পার্সিয়ান, জুলিয়ান) | মাঝারি | মাঝারি |
| ইনপুট ফরম্যাট নমনীয়তা | মাঝারি | সহজ |
| Struct রিটার্ন (গ্লোবালের বদলে) | নিম্ন | মাঝারি |
| Excel/Outlook-এ পোর্ট | মাঝারি | সহজ |
| চাঁদ-দেখা হিজরি API | নিম্ন | কঠিন |
| GUI UserForm | নিম্ন | মাঝারি |
| ইউনিট টেস্ট মডিউল | উচ্চ | সহজ |
| বঙ্গাব্দ লিপ ইয়ার সম্পূর্ণ বাস্তবায়ন | মাঝারি | মাঝারি |

---

## 15. Version History | সংস্করণ ইতিহাস

### English

| Version | Status | Key Changes |
|---------|--------|-------------|
| **v2.0** | ✅ Current | Added `dayNum_bn_pre` suffix variable for **all three** calendars (BanglaDate, EnglishDate, ArabicDate). Added `[BANG_DATE_NUM_PREFIX]` section to resource file. Introduced `arrBPrefixes(1..5)` array. Better UTF-8 file handling. Total variables increased from 57 → 63. |
| **v1.0** | 🗂 Legacy | Initial release. Multi-calendar support (Bengali, Gregorian, Hijri). Multi-language output (Bengali, English, Arabic). External resource file architecture. 57 global variables. |

### বাংলা — সংস্করণ ইতিহাস

| সংস্করণ | অবস্থা | মূল পরিবর্তন |
|---------|--------|-------------|
| **v2.0** | ✅ বর্তমান | তিনটি ক্যালেন্ডারেই `dayNum_bn_pre` প্রত্যয় ভেরিয়েবল। রিসোর্স ফাইলে `[BANG_DATE_NUM_PREFIX]` সেকশন। ৫৭ → ৬৩টি ভেরিয়েবল। |
| **v1.0** | 🗂 পুরানো | প্রাথমিক রিলিজ। তিন-ক্যালেন্ডার, তিন-ভাষা, বাহ্যিক রিসোর্স ফাইল। ৫৭ ভেরিয়েবল। |

---

## 16. Contributing | অবদান

### English

Contributions are very welcome! Here is how to participate:

**Reporting Bugs:**
Open a detailed issue including:
- Windows version and Word version
- The exact input date string that caused the problem
- The exact error message or incorrect output
- Contents of the first 5 lines of your resource file

**Submitting Code:**
1. Fork the repository
2. Create a branch: `git checkout -b feature/persian-calendar`
3. Follow the existing naming convention: `Abid_DEngine_[Calendar]_[Component]_[Script]`
4. Test against boundary dates (especially around April 13-15 for Bengali New Year)
5. Update this documentation to reflect changes
6. Submit a pull request with a clear description

**Naming Convention Rules:**
- New calendar variables: `Abid_DEngine_[NewCalendar]Date_[Component]_[Script]`
- New helper functions: `Private Function [DescriptiveName](...)`
- Resource file new sections: `[ALL_CAPS_UNDERSCORE]`
- Language index: extend sequentially (current: 1=bn, 2=en, 3=ar → next: 4=hi, 5=fa, etc.)

**Documentation:**
- All new features must be documented in **both English and Bangla**
- Add at least one usage example for every new function
- Update the variable reference tables

### বাংলা — অবদান করার নিয়ম

- বাগ রিপোর্টে Windows/Word সংস্করণ, ইনপুট তারিখ, ত্রুটি বার্তা অন্তর্ভুক্ত করুন
- নতুন ক্যালেন্ডার ভেরিয়েবলের জন্য বিদ্যমান নামকরণ নিয়ম অনুসরণ করুন
- সীমানা তারিখ পরীক্ষা করুন (বিশেষত এপ্রিল ১৩-১৫)
- সমস্ত নতুন বৈশিষ্ট্য বাংলা ও ইংরেজি উভয়েই ডকুমেন্ট করুন

---

## 17. License | লাইসেন্স

```
MIT License

Copyright (c) 2024 Abid Al Jamee

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

**In plain language | সহজ ভাষায়:**

| | English | বাংলা |
|---|---------|-------|
| ✅ | Use in personal projects | ব্যক্তিগত প্রজেক্টে ব্যবহার |
| ✅ | Use in commercial projects | বাণিজ্যিক প্রজেক্টে ব্যবহার |
| ✅ | Modify and distribute | পরিবর্তন ও বিতরণ |
| ✅ | Include in add-in packages | অ্যাড-ইন প্যাকেজে অন্তর্ভুক্ত |
| 👍 | Credit the original author | মূল লেখককে কৃতিত্ব দিন |
| 👍 | Include license file | লাইসেন্স ফাইল অন্তর্ভুক্ত করুন |
| ❌ | Hold author liable for damages | ক্ষতির জন্য লেখককে দায়ী করা |
| ❌ | Use author's name for endorsement | অনুমতি ছাড়া লেখকের নাম ব্যবহার |

---

## 18. Credits | কৃতিত্ব

### Author | লেখক
**Abid Al Jamee** — VBA Engine Developer, Bangladesh Bengali Calendar Implementation

[![Facebook](https://img.shields.io/badge/Facebook-abidaljamee-blue)](https://facebook.com/abidaljamee)

### Acknowledgments | কৃতজ্ঞতা
- **Government of Bangladesh / Bangla Academy** — Official Bengali calendar specifications
- **Microsoft VBA Team** — `vbCalHijri` calendar API
- **Open Source Community** — UTF-8 file handling techniques with ADODB.Stream

---

## Quick Reference Card | দ্রুত রেফারেন্স কার্ড

### The One Line You Always Need
```vba
Call ABID__DateENGINE(Format(Date, "DD/MM/YYYY"))
```

### Most-Used Variables at a Glance

| Want | Variable |
|------|----------|
| বাংলা তারিখ (প্রত্যয় সহ) | `Abid_DEngine_BanglaDate_dayNum_bn_pre` |
| বাংলা মাসের নাম | `Abid_DEngine_BanglaDate_monthName_bn` |
| বাংলা বছর | `Abid_DEngine_BanglaDate_yearNum_bn` |
| বাংলা যুগ | `Abid_DEngine_BanglaDate_Era_bn` |
| সপ্তাহের দিন (বাংলা) | `Abid_DEngine_BanglaDate_weekName_bn` |
| English day number | `Abid_DEngine_EnglishDate_dayNum_en` |
| English month name | `Abid_DEngine_EnglishDate_monthName_en` |
| English year | `Abid_DEngine_EnglishDate_yearNum_en` |
| Hijri day (Bengali script) | `Abid_DEngine_ArabicDate_dayNum_bn_pre` |
| Hijri month name (Arabic) | `Abid_DEngine_ArabicDate_monthName_ar` |
| Hijri year (Arabic script) | `Abid_DEngine_ArabicDate_yearNum_ar` |

### Standard Date String Recipes

```vba
' Full Bengali formal: সোমবার, ২রা বৈশাখ ১৪৩১ বঙ্গাব্দ
Abid_DEngine_BanglaDate_weekName_bn & ", " & Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & Abid_DEngine_BanglaDate_monthName_bn & " " & Abid_DEngine_BanglaDate_yearNum_bn & " " & Abid_DEngine_BanglaDate_Era_bn

' Full English: Monday, 15 April 2024 AD
Abid_DEngine_EnglishDate_weekName_en & ", " & Abid_DEngine_EnglishDate_dayNum_en & " " & Abid_DEngine_EnglishDate_monthName_en & " " & Abid_DEngine_EnglishDate_yearNum_en & " " & Abid_DEngine_EnglishDate_Era_en

' Full Hijri in Bengali: ৬ই শাওয়াল ১৪৪৫ হিজরি
Abid_DEngine_ArabicDate_dayNum_bn_pre & " " & Abid_DEngine_ArabicDate_monthName_bn & " " & Abid_DEngine_ArabicDate_yearNum_bn & " " & Abid_DEngine_ArabicDate_Era_bn

' Full Hijri in Arabic: ٦ شوال ١٤٤٥ هجري
Abid_DEngine_ArabicDate_dayNum_ar & " " & Abid_DEngine_ArabicDate_monthName_ar & " " & Abid_DEngine_ArabicDate_yearNum_ar & " " & Abid_DEngine_ArabicDate_Era_ar
```

---

*Last Updated: April 2026*
*Version: 2.0*
*Author: [Abid Al Jamee](https://facebook.com/abidaljamee)*
*License: MIT*

---

**End of Documentation | ডকুমেন্টেশন সমাপ্ত**
