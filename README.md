# ABID__DateENGINE v2.0 Documentation
# আবিদ_তারিখইঞ্জিন v2.0 ডকুমেন্টেশন

---

## Table of Contents | সূচিপত্র

**English:**
1. [Overview](#overview)
2. [Features](#features)
3. [System Requirements](#system-requirements)
4. [Installation](#installation)
5. [Resource File Structure](#resource-file-structure)
6. [Global Variables Reference](#global-variables-reference)
7. [Function Reference](#function-reference)
8. [Usage Examples](#usage-examples)
9. [Calendar Conversion Logic](#calendar-conversion-logic)
10. [Troubleshooting](#troubleshooting)
11. [License](#license)

**বাংলা:**
1. [সংক্ষিপ্ত বিবরণ](#সংক্ষিপ্ত-বিবরণ)
2. [বৈশিষ্ট্যসমূহ](#বৈশিষ্ট্যসমূহ)
3. [সিস্টেম প্রয়োজনীয়তা](#সিস্টেম-প্রয়োজনীয়তা)
4. [ইনস্টলেশন](#ইনস্টলেশন)
5. [রিসোর্স ফাইল কাঠামো](#রিসোর্স-ফাইল-কাঠামো)
6. [গ্লোবাল ভেরিয়েবল রেফারেন্স](#গ্লোবাল-ভেরিয়েবল-রেফারেন্স)
7. [ফাংশন রেফারেন্স](#ফাংশন-রেফারেন্স)
8. [ব্যবহারের উদাহরণ](#ব্যবহারের-উদাহরণ)
9. [ক্যালেন্ডার রূপান্তর যুক্তি](#ক্যালেন্ডার-রূপান্তর-যুক্তি)
10. [সমস্যা সমাধান](#সমস্যা-সমাধান)
11. [লাইসেন্স](#লাইসেন্স)

---

## Overview

**ABID__DateENGINE** is a powerful Microsoft Word VBA macro engine that provides comprehensive multi-calendar date conversion capabilities. It simultaneously converts dates between three calendar systems (Bengali, Gregorian, and Hijri) and provides localized output in three languages (Bengali, English, and Arabic).

### Key Capabilities:
- **Multi-Calendar Support**: Bangla Calendar (বাংলা সন), Gregorian Calendar, Hijri Calendar (হিজরি)
- **Multi-Language Output**: Bengali (বাংলা), English, Arabic (العربية)
- **Resource-Driven Architecture**: External UTF-8 resource file for easy customization
- **Global Variable System**: 60 public variables accessible from any VBA macro
- **Automatic Date Suffix**: Intelligent Bengali date suffix system (১লা, ২রা, ৪ঠা, ৫ই, ১৯শে)

---

## সংক্ষিপ্ত বিবরণ

**আবিদ_তারিখইঞ্জিন** হলো একটি শক্তিশালী Microsoft Word VBA ম্যাক্রো ইঞ্জিন যা ব্যাপক বহু-ক্যালেন্ডার তারিখ রূপান্তর সুবিধা প্রদান করে। এটি একসাথে তিনটি ক্যালেন্ডার সিস্টেম (বাংলা, গ্রেগরিয়ান এবং হিজরি) এর মধ্যে তারিখ রূপান্তর করে এবং তিনটি ভাষায় (বাংলা, ইংরেজি এবং আরবি) স্থানীয়করণ আউটপুট প্রদান করে।

### মূল সক্ষমতা:
- **বহু-ক্যালেন্ডার সমর্থন**: বাংলা ক্যালেন্ডার (বাংলা সন), গ্রেগরিয়ান ক্যালেন্ডার, হিজরি ক্যালেন্ডার
- **বহু-ভাষা আউটপুট**: বাংলা, ইংরেজি, আরবি (العربية)
- **রিসোর্স-চালিত আর্কিটেকচার**: সহজ কাস্টমাইজেশনের জন্য বাহ্যিক UTF-8 রিসোর্স ফাইল
- **গ্লোবাল ভেরিয়েবল সিস্টেম**: যেকোনো VBA ম্যাক্রো থেকে অ্যাক্সেসযোগ্য ৬০টি পাবলিক ভেরিয়েবল
- **স্বয়ংক্রিয় তারিখ প্রত্যয়**: বুদ্ধিমান বাংলা তারিখ প্রত্যয় সিস্টেম (১লা, ২রা, ৪ঠা, ৫ই, ১৯শে)

---

## Features

### 1. **Triple Calendar Conversion**
The engine converts a single input date into three calendar systems:
- **Bengali Calendar (Bangabda)**: Official Bangladesh calendar following government regulations
- **Gregorian Calendar**: International standard calendar (AD/CE)
- **Hijri Calendar**: Islamic lunar calendar (AH)

### 2. **Triple Language Localization**
Each date component is available in three language scripts:
- **Bengali Script** (bn): ০১, বৈশাখ, শনিবার
- **English Script** (en): 01, Baishakh, Saturday
- **Arabic Script** (ar): ٠١, محرم, السبت

### 3. **Comprehensive Date Components**
For each calendar, the engine provides:
- Day number (with Bengali suffix support)
- Month number
- Year number
- Day of week name
- Month name
- Era designation

### 4. **Smart Bengali Date Suffix**
Automatically applies correct Bengali ordinal suffix:
- **১লা** (1st day)
- **২রা, ৩রা** (2nd, 3rd days)
- **৪ঠা** (4th day)
- **৫ই to ১৮ই** (5th to 18th days)
- **১৯শে to ৩১শে** (19th to 31st days)

---

## বৈশিষ্ট্যসমূহ

### ১. **ত্রিগুণ ক্যালেন্ডার রূপান্তর**
ইঞ্জিনটি একটি একক ইনপুট তারিখকে তিনটি ক্যালেন্ডার সিস্টেমে রূপান্তরিত করে:
- **বাংলা ক্যালেন্ডার (বঙ্গাব্দ)**: সরকারি নিয়মাবলী অনুসরণকারী অফিশিয়াল বাংলাদেশ ক্যালেন্ডার
- **গ্রেগরিয়ান ক্যালেন্ডার**: আন্তর্জাতিক মান ক্যালেন্ডার (খ্রিস্টাব্দ)
- **হিজরি ক্যালেন্ডার**: ইসলামিক চন্দ্র ক্যালেন্ডার (হিজরি)

### ২. **ত্রিগুণ ভাষা স্থানীয়করণ**
প্রতিটি তারিখ উপাদান তিনটি ভাষা লিপিতে উপলব্ধ:
- **বাংলা লিপি** (bn): ০১, বৈশাখ, শনিবার
- **ইংরেজি লিপি** (en): 01, Baishakh, Saturday
- **আরবি লিপি** (ar): ٠١, محرم, السبت

### ৩. **ব্যাপক তারিখ উপাদান**
প্রতিটি ক্যালেন্ডারের জন্য, ইঞ্জিন প্রদান করে:
- দিন নম্বর (বাংলা প্রত্যয় সমর্থন সহ)
- মাস নম্বর
- বছর নম্বর
- সপ্তাহের দিনের নাম
- মাসের নাম
- যুগ পদবী

### ৪. **স্মার্ট বাংলা তারিখ প্রত্যয়**
স্বয়ংক্রিয়ভাবে সঠিক বাংলা ক্রমবাচক প্রত্যয় প্রয়োগ করে:
- **১লা** (১ম দিন)
- **২রা, ৩রা** (২য়, ৩য় দিন)
- **৪ঠা** (৪র্থ দিন)
- **৫ই থেকে ১৮ই** (৫ম থেকে ১৮তম দিন)
- **১৯শে থেকে ৩১শে** (১৯তম থেকে ৩১তম দিন)

---

## System Requirements

### Minimum Requirements:
- **Microsoft Word**: 2007 or later (2010, 2013, 2016, 2019, 2021, 365)
- **Operating System**: Windows 7 or later
- **VBA Support**: Must be enabled in Word
- **File System Access**: Write access to %APPDATA% directory
- **Required Libraries**: 
  - Microsoft Scripting Runtime (FSO)
  - ActiveX Data Objects (ADODB) for UTF-8 file reading

### Dependencies:
1. **Resource File**: `ABID_DateENGINE_resources.txt` (UTF-8 encoded)
2. **File Location**: `%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\`

---

## সিস্টেম প্রয়োজনীয়তা

### ন্যূনতম প্রয়োজনীয়তা:
- **Microsoft Word**: ২০০৭ বা পরবর্তী (২০১০, ২০১৩, ২০১৬, ২০১৯, ২০২১, ৩৬৫)
- **অপারেটিং সিস্টেম**: Windows ৭ বা পরবর্তী
- **VBA সমর্থন**: Word-এ সক্ষম থাকতে হবে
- **ফাইল সিস্টেম অ্যাক্সেস**: %APPDATA% ডিরেক্টরিতে লেখার অ্যাক্সেস
- **প্রয়োজনীয় লাইব্রেরি**: 
  - Microsoft Scripting Runtime (FSO)
  - ActiveX Data Objects (ADODB) UTF-8 ফাইল পড়ার জন্য

### নির্ভরতা:
১. **রিসোর্স ফাইল**: `ABID_DateENGINE_resources.txt` (UTF-8 এনকোডেড)
২. **ফাইল অবস্থান**: `%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\`

---

## Installation

### Step 1: Import the VBA Module
1. Open Microsoft Word
2. Press `Alt + F11` to open VBA Editor
3. Go to **File** → **Import File**
4. Select `ABID__DateENGINE_v2.bas`
5. The module will appear in your VBA project

### Step 2: Create Resource Directory
Create the following directory structure:
```
%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

**Windows Path Example:**
```
C:\Users\YourUsername\AppData\Roaming\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

### Step 3: Create Resource File
Create a file named `ABID_DateENGINE_resources.txt` in the above directory with UTF-8 encoding.

---

## ইনস্টলেশন

### ধাপ ১: VBA মডিউল ইম্পোর্ট করুন
১. Microsoft Word খুলুন
২. VBA এডিটর খুলতে `Alt + F11` চাপুন
৩. **File** → **Import File** এ যান
৪. `ABID__DateENGINE_v2.bas` নির্বাচন করুন
৫. মডিউলটি আপনার VBA প্রজেক্টে প্রদর্শিত হবে

### ধাপ ২: রিসোর্স ডিরেক্টরি তৈরি করুন
নিম্নলিখিত ডিরেক্টরি কাঠামো তৈরি করুন:
```
%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

**Windows পথের উদাহরণ:**
```
C:\Users\YourUsername\AppData\Roaming\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\
```

### ধাপ ৩: রিসোর্স ফাইল তৈরি করুন
উপরের ডিরেক্টরিতে UTF-8 এনকোডিং সহ `ABID_DateENGINE_resources.txt` নামে একটি ফাইল তৈরি করুন।

---

## Resource File Structure

The resource file must be UTF-8 encoded and follow this exact structure:

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
পৌষ|Poush|পৌষ
মাঘ|Magh|مাغ
ফাল্গুন|Falgun|فالغون
চৈত্র|Choitro|চৈত্র

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

### Format Rules:
- Each section begins with `[SECTION_NAME]`
- Data rows use pipe delimiter: `Bengali|English|Arabic`
- Prefix row uses comma delimiter: `suffix1,suffix2,suffix3,suffix4,suffix5`
- File must be saved as UTF-8 encoding
- Line endings can be CRLF or LF (automatically handled)

---

## রিসোর্স ফাইল কাঠামো

রিসোর্স ফাইলটি অবশ্যই UTF-8 এনকোডেড হতে হবে এবং এই সঠিক কাঠামো অনুসরণ করতে হবে:

*(উপরের ইংরেজি বিভাগে দেখানো একই কাঠামো)*

### ফরম্যাট নিয়ম:
- প্রতিটি বিভাগ `[SECTION_NAME]` দিয়ে শুরু হয়
- ডেটা সারিগুলি পাইপ ডিলিমিটার ব্যবহার করে: `বাংলা|ইংরেজি|আরবি`
- প্রিফিক্স সারি কমা ডিলিমিটার ব্যবহার করে: `প্রত্যয়১,প্রত্যয়২,প্রত্যয়৩,প্রত্যয়৪,প্রত্যয়৫`
- ফাইল UTF-8 এনকোডিং হিসাবে সংরক্ষণ করতে হবে
- লাইন এন্ডিং CRLF বা LF হতে পারে (স্বয়ংক্রিয়ভাবে পরিচালনা করা হয়)

---

## Global Variables Reference

The engine populates **60 public variables** that can be accessed from any VBA macro in your document. All variables follow the naming convention: `Abid_DEngine_[Calendar]Date_[Component]_[Language]`

### Variable Naming Pattern:
```
Abid_DEngine_[BanglaDate|EnglishDate|ArabicDate]_[Component]_[bn|en|ar]
```

### Complete Variable List:

#### **1. Day Numbers (তারিখ নম্বর)**

| Variable Name | Type | Description | Example Value |
|--------------|------|-------------|---------------|
| `Abid_DEngine_BanglaDate_dayNum_bn` | String | Bengali date day in Bengali numerals | ০১ |
| `Abid_DEngine_BanglaDate_dayNum_bn_pre` | String | Bengali date day with suffix | ১লা, ৫ই, ১৯শে |
| `Abid_DEngine_BanglaDate_monthNum_en` | String | Bengali date month in English numerals | 01 |
| `Abid_DEngine_BanglaDate_yearNum_ar` | String | Bengali date year in Arabic numerals | ١٤٣١ |
| `Abid_DEngine_EnglishDate_dayNum_bn` | String | Gregorian date day in Bengali numerals | ১৫ |
| `Abid_DEngine_EnglishDate_monthNum_en` | String | Gregorian date month in English numerals | 04 |
| `Abid_DEngine_EnglishDate_yearNum_ar` | String | Gregorian date year in Arabic numerals | ٢٠٢٤ |
| `Abid_DEngine_ArabicDate_dayNum_bn` | String | Hijri date day in Bengali numerals | ০৮ |
| `Abid_DEngine_ArabicDate_monthNum_en` | String | Hijri date month in English numerals | 10 |
| `Abid_DEngine_ArabicDate_yearNum_ar` | String | Hijri date year in Arabic numerals | ١٤٤٥ |

#### **2. Week Names (সপ্তাহের দিনের নাম)**

| Variable Name | Type | Description | Example Value |
|--------------|------|-------------|---------------|
| `Abid_DEngine_BanglaDate_weekName_bn` | String | Weekday name in Bengali | শনিবার |
| `Abid_DEngine_BanglaDate_weekName_en` | String | Weekday name in English | Saturday |
| `Abid_DEngine_BanglaDate_weekName_ar` | String | Weekday name in Arabic | السبت |
| `Abid_DEngine_EnglishDate_weekName_bn` | String | Weekday name in Bengali | শনিবার |
| `Abid_DEngine_EnglishDate_weekName_en` | String | Weekday name in English | Saturday |
| `Abid_DEngine_EnglishDate_weekName_ar` | String | Weekday name in Arabic | السبت |
| `Abid_DEngine_ArabicDate_weekName_bn` | String | Weekday name in Bengali | শনিবার |
| `Abid_DEngine_ArabicDate_weekName_en` | String | Weekday name in English | Saturday |
| `Abid_DEngine_ArabicDate_weekName_ar` | String | Weekday name in Arabic | السبت |

#### **3. Month Names (মাসের নাম)**

| Variable Name | Type | Description | Example Value |
|--------------|------|-------------|---------------|
| `Abid_DEngine_BanglaDate_monthName_bn` | String | Bengali month name in Bengali | বৈশাখ |
| `Abid_DEngine_BanglaDate_monthName_en` | String | Bengali month name in English | Baishakh |
| `Abid_DEngine_BanglaDate_monthName_ar` | String | Bengali month name in Arabic | بيشاخ |
| `Abid_DEngine_EnglishDate_monthName_bn` | String | Gregorian month name in Bengali | এপ্রিল |
| `Abid_DEngine_EnglishDate_monthName_en` | String | Gregorian month name in English | April |
| `Abid_DEngine_EnglishDate_monthName_ar` | String | Gregorian month name in Arabic | ابريل |
| `Abid_DEngine_ArabicDate_monthName_bn` | String | Hijri month name in Bengali | শাওয়াল |
| `Abid_DEngine_ArabicDate_monthName_en` | String | Hijri month name in English | Shawwal |
| `Abid_DEngine_ArabicDate_monthName_ar` | String | Hijri month name in Arabic | شوال |

#### **4. Era Designations (যুগ পদবী)**

| Variable Name | Type | Description | Example Value |
|--------------|------|-------------|---------------|
| `Abid_DEngine_BanglaDate_Era_bn` | String | Era name in Bengali | বঙ্গাব্দ |
| `Abid_DEngine_BanglaDate_Era_en` | String | Era name in English | Bangabda |
| `Abid_DEngine_BanglaDate_Era_ar` | String | Era name in Arabic | بنغابدا |
| `Abid_DEngine_EnglishDate_Era_bn` | String | Era name in Bengali | খ্রিস্টাব্দ |
| `Abid_DEngine_EnglishDate_Era_en` | String | Era name in English | AD |
| `Abid_DEngine_EnglishDate_Era_ar` | String | Era name in Arabic | ميلادي |
| `Abid_DEngine_ArabicDate_Era_bn` | String | Era name in Bengali | হিজরি |
| `Abid_DEngine_ArabicDate_Era_en` | String | Era name in English | AH |
| `Abid_DEngine_ArabicDate_Era_ar` | String | Era name in Arabic | هجري |

---

## গ্লোবাল ভেরিয়েবল রেফারেন্স

ইঞ্জিনটি **৬০টি পাবলিক ভেরিয়েবল** পূর্ণ করে যা আপনার ডকুমেন্টের যেকোনো VBA ম্যাক্রো থেকে অ্যাক্সেস করা যায়। সকল ভেরিয়েবল এই নামকরণ নিয়ম অনুসরণ করে: `Abid_DEngine_[ক্যালেন্ডার]Date_[উপাদান]_[ভাষা]`

*(উপরের ইংরেজি বিভাগে সম্পূর্ণ টেবিল দেখুন)*

---

## Function Reference

### Main Function: `ABID__DateENGINE`

```vba
Public Sub ABID__DateENGINE(ByVal InputDateStr As String)
```

**Purpose:** Converts a date from DD/MM/YYYY format into three calendar systems and populates all global variables.

**Parameters:**
- `InputDateStr` (String): Date in DD/MM/YYYY format (e.g., "15/04/2024")

**Returns:** None (populates global variables)

**Throws:**
- Error message if date format is invalid
- Error message if resource file is not found

**Example:**
```vba
Call ABID__DateENGINE("15/04/2024")
' Now all 60 global variables are populated and ready to use
```

---

### Helper Function: `LoadDateResources` (Private)

```vba
Private Function LoadDateResources(ByRef arrBMonths() As String, _
                                   ByRef arrGMonths() As String, _
                                   ByRef arrHMonths() As String, _
                                   ByRef arrEras() As String, _
                                   ByRef arrWeeks() As String, _
                                   ByRef arrNums() As String, _
                                   ByRef arrBPrefixes() As String) As Boolean
```

**Purpose:** Loads localized data from the external UTF-8 resource file.

**Parameters:** Seven array references passed by reference
**Returns:** `True` if successful, `False` if file not found

**Resource Arrays:**
- `arrBMonths`: Bangla month names [1-12, 1-3]
- `arrGMonths`: Gregorian month names [1-12, 1-3]
- `arrHMonths`: Hijri month names [1-12, 1-3]
- `arrEras`: Era designations [1-3, 1-3]
- `arrWeeks`: Weekday names [1-7, 1-3]
- `arrNums`: Numerals [0-9, 1-3]
- `arrBPrefixes`: Bengali suffixes [1-5]

---

### Helper Function: `ToLocalizedNum` (Private)

```vba
Private Function ToLocalizedNum(ByVal strNum As String, _
                                ByVal langIdx As Integer, _
                                ByRef arrNums() As String) As String
```

**Purpose:** Converts English numerals to the target language numerals.

**Parameters:**
- `strNum`: Number string in English (e.g., "15")
- `langIdx`: Language index (1=Bengali, 2=English, 3=Arabic)
- `arrNums`: Numeral mapping array

**Returns:** Localized number string (e.g., "১৫" for Bengali)

---

## ফাংশন রেফারেন্স

### প্রধান ফাংশন: `ABID__DateENGINE`

```vba
Public Sub ABID__DateENGINE(ByVal InputDateStr As String)
```

**উদ্দেশ্য:** DD/MM/YYYY ফরম্যাট থেকে একটি তারিখকে তিনটি ক্যালেন্ডার সিস্টেমে রূপান্তরিত করে এবং সমস্ত গ্লোবাল ভেরিয়েবল পূর্ণ করে।

**প্যারামিটার:**
- `InputDateStr` (স্ট্রিং): DD/MM/YYYY ফরম্যাটে তারিখ (যেমন, "১৫/০৪/২০২৪")

**রিটার্ন:** কিছুই নয় (গ্লোবাল ভেরিয়েবল পূর্ণ করে)

**ত্রুটি:**
- তারিখ ফরম্যাট অবৈধ হলে ত্রুটি বার্তা
- রিসোর্স ফাইল না পাওয়া গেলে ত্রুটি বার্তা

**উদাহরণ:**
```vba
Call ABID__DateENGINE("15/04/2024")
' এখন সমস্ত ৬০টি গ্লোবাল ভেরিয়েবল পূর্ণ এবং ব্যবহারের জন্য প্রস্তুত
```

---

## Usage Examples

### Example 1: Basic Date Conversion

```vba
Sub Example_BasicDateConversion()
    ' Convert today's date
    Dim todayStr As String
    todayStr = Format(Date, "DD/MM/YYYY")
    
    ' Call the engine
    Call ABID__DateENGINE(todayStr)
    
    ' Access Bengali date
    Debug.Print "Bengali Date: " & Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                Abid_DEngine_BanglaDate_monthName_bn & " " & _
                Abid_DEngine_BanglaDate_yearNum_ar & " " & _
                Abid_DEngine_BanglaDate_Era_bn
    ' Output: ১লা বৈশাখ ১৪৩১ বঙ্গাব্দ
End Sub
```

### Example 2: Insert Formatted Date in Document

```vba
Sub Example_InsertBengaliDate()
    ' Specific date
    Call ABID__DateENGINE("15/04/2024")
    
    ' Build formatted string
    Dim bengaliDate As String
    bengaliDate = Abid_DEngine_BanglaDate_weekName_bn & ", " & _
                  Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                  Abid_DEngine_BanglaDate_monthName_bn & " " & _
                  Abid_DEngine_BanglaDate_yearNum_ar & " " & _
                  Abid_DEngine_BanglaDate_Era_bn
    
    ' Insert at cursor position
    Selection.TypeText bengaliDate
    ' Output: সোমবার, ১লা বৈশাখ ১৪৩১ বঙ্গাব্দ
End Sub
```

### Example 3: Multi-Calendar Display

```vba
Sub Example_MultiCalendarDisplay()
    Call ABID__DateENGINE("15/04/2024")
    
    Dim output As String
    
    ' Bengali Date
    output = "বাংলা: " & Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
             Abid_DEngine_BanglaDate_monthName_bn & " " & _
             Abid_DEngine_BanglaDate_yearNum_ar & vbCrLf
    
    ' English Date
    output = output & "English: " & Abid_DEngine_EnglishDate_monthName_en & " " & _
             Abid_DEngine_EnglishDate_dayNum_bn & ", " & _
             Abid_DEngine_EnglishDate_yearNum_ar & " " & _
             Abid_DEngine_EnglishDate_Era_en & vbCrLf
    
    ' Arabic Date
    output = output & "Arabic: " & Abid_DEngine_ArabicDate_dayNum_bn & " " & _
             Abid_DEngine_ArabicDate_monthName_ar & " " & _
             Abid_DEngine_ArabicDate_yearNum_ar & " " & _
             Abid_DEngine_ArabicDate_Era_ar
    
    MsgBox output
End Sub
```

### Example 4: Table Population

```vba
Sub Example_PopulateDateTable()
    Call ABID__DateENGINE("15/04/2024")
    
    ' Assume table exists at cursor with 3 columns
    Dim tbl As Table
    Set tbl = Selection.Tables(1)
    
    ' Row 1: Bengali
    tbl.Cell(1, 1).Range.Text = "বাংলা তারিখ"
    tbl.Cell(1, 2).Range.Text = Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                                Abid_DEngine_BanglaDate_monthName_bn & " " & _
                                Abid_DEngine_BanglaDate_yearNum_ar
    
    ' Row 2: English
    tbl.Cell(2, 1).Range.Text = "English Date"
    tbl.Cell(2, 2).Range.Text = Abid_DEngine_EnglishDate_monthName_en & " " & _
                                Abid_DEngine_EnglishDate_dayNum_bn & ", " & _
                                Abid_DEngine_EnglishDate_yearNum_ar
    
    ' Row 3: Arabic
    tbl.Cell(3, 1).Range.Text = "التاريخ العربي"
    tbl.Cell(3, 2).Range.Text = Abid_DEngine_ArabicDate_dayNum_bn & " " & _
                                Abid_DEngine_ArabicDate_monthName_ar & " " & _
                                Abid_DEngine_ArabicDate_yearNum_ar
End Sub
```

### Example 5: Custom Date Formatter Function

```vba
Function FormatCustomDate(dateStr As String, calendar As String, language As String) As String
    Call ABID__DateENGINE(dateStr)
    
    Dim result As String
    
    Select Case calendar
        Case "Bangla"
            Select Case language
                Case "bn"
                    result = Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                            Abid_DEngine_BanglaDate_monthName_bn & " " & _
                            Abid_DEngine_BanglaDate_yearNum_ar
                Case "en"
                    result = Abid_DEngine_BanglaDate_monthName_en & " " & _
                            Abid_DEngine_BanglaDate_dayNum_bn & ", " & _
                            Abid_DEngine_BanglaDate_yearNum_ar
            End Select
        Case "English"
            Select Case language
                Case "bn"
                    result = Abid_DEngine_EnglishDate_dayNum_bn & " " & _
                            Abid_DEngine_EnglishDate_monthName_bn & " " & _
                            Abid_DEngine_EnglishDate_yearNum_ar
                Case "en"
                    result = Abid_DEngine_EnglishDate_monthName_en & " " & _
                            Abid_DEngine_EnglishDate_dayNum_bn & ", " & _
                            Abid_DEngine_EnglishDate_yearNum_ar
            End Select
        Case "Arabic"
            Select Case language
                Case "ar"
                    result = Abid_DEngine_ArabicDate_dayNum_bn & " " & _
                            Abid_DEngine_ArabicDate_monthName_ar & " " & _
                            Abid_DEngine_ArabicDate_yearNum_ar
            End Select
    End Select
    
    FormatCustomDate = result
End Function

' Usage:
' MsgBox FormatCustomDate("15/04/2024", "Bangla", "bn")
```

---

## ব্যবহারের উদাহরণ

### উদাহরণ ১: মৌলিক তারিখ রূপান্তর

```vba
Sub Example_BasicDateConversion()
    ' আজকের তারিখ রূপান্তর করুন
    Dim todayStr As String
    todayStr = Format(Date, "DD/MM/YYYY")
    
    ' ইঞ্জিন কল করুন
    Call ABID__DateENGINE(todayStr)
    
    ' বাংলা তারিখ অ্যাক্সেস করুন
    Debug.Print "Bengali Date: " & Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                Abid_DEngine_BanglaDate_monthName_bn & " " & _
                Abid_DEngine_BanglaDate_yearNum_ar & " " & _
                Abid_DEngine_BanglaDate_Era_bn
    ' আউটপুট: ১লা বৈশাখ ১৪৩১ বঙ্গাব্দ
End Sub
```

*(অন্যান্য উদাহরণগুলি উপরের ইংরেজি বিভাগে দেখুন)*

---

## Calendar Conversion Logic

### Bengali Calendar (বাংলা ক্যালেন্ডার)

The engine implements the **official Bangladesh Bengali calendar** system following government regulations:

#### Year Calculation:
- If date is before April 14: `BengaliYear = GregorianYear - 594`
- If date is April 14 or after: `BengaliYear = GregorianYear - 593`

#### Month & Day Mapping:
The conversion follows the fixed government calendar where:
- First 5 months (Baishakh to Bhadro): 31 days each
- Next 6 months (Ashwin to Falgun): 30 days each
- Last month (Choitro): 30 days (31 in leap year)

**Key Dates:**
- Bengali New Year (১লা বৈশাখ) = April 14 (Gregorian)
- Month transitions occur on the 14th, 15th, 16th, or 17th of Gregorian months

#### Leap Year Handling:
Leap years in the Gregorian calendar affect the calculation for dates in Falgun (ফাল্গুন), adding one extra day to account for February 29th.

### Date Suffix Rules (তারিখ প্রত্যয় নিয়ম):

| Day | Suffix | Example |
|-----|--------|---------|
| 1 | লা | ১লা |
| 2-3 | রা | ২রা, ৩রা |
| 4 | ঠা | ৪ঠা |
| 5-18 | ই | ৫ই, ১০ই, ১৮ই |
| 19-31 | শে | ১৯শে, ২৫শে, ৩০শে |

---

### Hijri Calendar (হিজরি ক্যালেন্ডার)

The engine uses VBA's built-in Hijri calendar conversion:

```vba
' Switch to Hijri calendar
VBA.Calendar = vbCalHijri
hDay = Day(dt)
hMonth = Month(dt)
hYear = Year(dt)
' Switch back to Gregorian
VBA.Calendar = vbCalGreg
```

**Note:** The Hijri calendar is a lunar calendar with months of 29 or 30 days. The exact conversion may vary slightly depending on moon sighting methodologies.

---

### Gregorian Calendar (গ্রেগরিয়ান ক্যালেন্ডার)

Standard international calendar. VBA's native `Date` object uses Gregorian by default.

---

## ক্যালেন্ডার রূপান্তর যুক্তি

*(উপরের ইংরেজি বিভাগে বিস্তারিত ব্যাখ্যা দেখুন)*

### বাংলা ক্যালেন্ডার

ইঞ্জিনটি সরকারি নিয়মাবলী অনুসরণ করে **অফিশিয়াল বাংলাদেশ বাংলা ক্যালেন্ডার** সিস্টেম বাস্তবায়ন করে:

#### বছর গণনা:
- তারিখ ১৪ এপ্রিলের আগে হলে: `বাংলা বছর = গ্রেগরিয়ান বছর - ৫৯৪`
- তারিখ ১৪ এপ্রিল বা পরে হলে: `বাংলা বছর = গ্রেগরিয়ান বছর - ৫৯৩`

#### মাস এবং দিন ম্যাপিং:
রূপান্তরটি নির্ধারিত সরকারি ক্যালেন্ডার অনুসরণ করে যেখানে:
- প্রথম ৫ মাস (বৈশাখ থেকে ভাদ্র): প্রতিটি ৩১ দিন
- পরবর্তী ৬ মাস (আশ্বিন থেকে ফাল্গুন): প্রতিটি ৩০ দিন
- শেষ মাস (চৈত্র): ৩০ দিন (লিপ ইয়ারে ৩১ দিন)

**মূল তারিখসমূহ:**
- বাংলা নববর্ষ (১লা বৈশাখ) = ১৪ এপ্রিল (গ্রেগরিয়ান)
- মাস পরিবর্তন গ্রেগরিয়ান মাসের ১৪, ১৫, ১৬ বা ১৭ তারিখে ঘটে

---

## Troubleshooting

### Common Issues and Solutions:

#### Issue 1: "Resource file not found" Error

**Cause:** The resource file is missing or in wrong location

**Solution:**
1. Verify the file exists at: `%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\ABID_DateENGINE_resources.txt`
2. Check exact spelling and case sensitivity
3. Ensure you have read permissions for the directory

#### Issue 2: Bengali Characters Display as Boxes

**Cause:** Resource file is not UTF-8 encoded

**Solution:**
1. Open the resource file in Notepad++
2. Go to **Encoding** → **Convert to UTF-8**
3. Save the file
4. Restart Word and try again

#### Issue 3: Wrong Bengali Date Calculated

**Cause:** Input date format is incorrect

**Solution:**
- Ensure date is in DD/MM/YYYY format
- Use leading zeros: "01/04/2024" not "1/4/2024"
- Month must be between 01-12
- Day must be valid for the month

#### Issue 4: Variables are Empty After Running

**Cause:** Function exited early due to error

**Solution:**
1. Check if error message was displayed
2. Verify resource file is complete with all sections
3. Ensure all pipe delimiters (|) are present in resource file
4. Check Debug window (Ctrl+G in VBA Editor) for error messages

#### Issue 5: ADODB Error on File Reading

**Cause:** ActiveX Data Objects library reference missing

**Solution:**
1. In VBA Editor, go to **Tools** → **References**
2. Check "Microsoft ActiveX Data Objects 6.1 Library" (or latest version)
3. Click OK and try again

---

## সমস্যা সমাধান

### সাধারণ সমস্যা এবং সমাধান:

#### সমস্যা ১: "Resource file not found" ত্রুটি

**কারণ:** রিসোর্স ফাইল অনুপস্থিত বা ভুল অবস্থানে

**সমাধান:**
১. ফাইলটি এই অবস্থানে আছে কিনা যাচাই করুন: `%APPDATA%\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\ABID_DateENGINE_resources.txt`
২. সঠিক বানান এবং কেস সেনসিটিভিটি চেক করুন
৩. ডিরেক্টরিতে পড়ার অনুমতি আছে কিনা নিশ্চিত করুন

#### সমস্যা ২: বাংলা অক্ষর বক্স হিসাবে প্রদর্শিত হয়

**কারণ:** রিসোর্স ফাইল UTF-8 এনকোডেড নয়

**সমাধান:**
১. Notepad++ এ রিসোর্স ফাইল খুলুন
২. **Encoding** → **Convert to UTF-8** এ যান
৩. ফাইল সংরক্ষণ করুন
৪. Word পুনরায় চালু করুন এবং আবার চেষ্টা করুন

*(অন্যান্য সমস্যা সমাধানের জন্য উপরের ইংরেজি বিভাগ দেখুন)*

---

## Performance Considerations

### Optimization Tips:

1. **Resource File Caching**: The resource file is loaded on each function call. For batch processing, consider loading once and reusing arrays.

2. **Variable Access**: Global variables are faster to access than repeated function calls. Call `ABID__DateENGINE` once per date, then use variables multiple times.

3. **Date Format**: Always provide dates in DD/MM/YYYY format to avoid conversion overhead.

4. **Batch Operations**: If processing multiple dates, create a wrapper function that caches resources.

---

## Best Practices

### 1. Error Handling in Your Macros

```vba
Sub SafeDateConversion()
    On Error GoTo ErrorHandler
    
    Call ABID__DateENGINE("15/04/2024")
    ' Use variables here
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Date conversion failed: " & Err.Description, vbCritical
End Sub
```

### 2. Input Validation

```vba
Function IsValidDateFormat(dateStr As String) As Boolean
    Dim parts() As String
    parts = Split(dateStr, "/")
    
    If UBound(parts) <> 2 Then
        IsValidDateFormat = False
        Exit Function
    End If
    
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        IsValidDateFormat = False
        Exit Function
    End If
    
    IsValidDateFormat = True
End Function
```

### 3. Centralized Date Formatting

Create a module with standard date formatting functions:

```vba
' Standard Bengali date format
Function GetStandardBengaliDate() As String
    GetStandardBengaliDate = Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
                            Abid_DEngine_BanglaDate_monthName_bn & " " & _
                            Abid_DEngine_BanglaDate_yearNum_ar & " " & _
                            Abid_DEngine_BanglaDate_Era_bn
End Function

' Standard English date format
Function GetStandardEnglishDate() As String
    GetStandardEnglishDate = Abid_DEngine_EnglishDate_monthName_en & " " & _
                            Abid_DEngine_EnglishDate_dayNum_bn & ", " & _
                            Abid_DEngine_EnglishDate_yearNum_ar & " " & _
                            Abid_DEngine_EnglishDate_Era_en
End Function
```

---

## Advanced Usage

### Custom Resource File Location

If you want to use a different resource file location, modify the `LoadDateResources` function:

```vba
' Original line 249:
filePath = Environ("AppData") & "\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\ABID_DateENGINE_resources.txt"

' Change to custom path:
filePath = "C:\MyCustomPath\ABID_DateENGINE_resources.txt"
```

### Adding New Languages

To add a fourth language column to the resource file:

1. Modify resource file: Add `|Language4` to each row
2. Modify array dimensions: Change `ReDim arrBMonths(1 To 12, 1 To 4)`
3. Add new global variables with `_l4` suffix
4. Update parsing logic to handle index 4

---

## Version History

### Version 2 (v2) - Current
- Added Bengali date suffix support (`dayNum_bn_pre`)
- Implemented `[BANG_DATE_NUM_PREFIX]` section in resource file
- Smart suffix logic: লা, রা, ঠা, ই, শে
- Enhanced resource loading with prefix array
- Better UTF-8 file handling

### Version 1 (v1)
- Initial release
- Multi-calendar support (Bengali, Gregorian, Hijri)
- Multi-language output (Bengali, English, Arabic)
- External resource file architecture
- 57 global variables (before suffix addition)

---

## Contributing

This is an open-source project. Contributions are welcome!

### How to Contribute:

1. **Report Issues**: Submit bug reports with detailed error messages
2. **Suggest Features**: Propose new calendar systems or language support
3. **Improve Documentation**: Help translate or clarify documentation
4. **Submit Code**: Fork, modify, and submit pull requests

### Guidelines:

- Follow the existing naming convention: `Abid_DEngine_[Component]_[Language]`
- Document all new variables and functions
- Test with multiple date ranges
- Ensure backward compatibility
- Update resource file format if needed

---

## License

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

### Usage Rights:

✅ **You CAN:**
- Use this engine in personal projects
- Use this engine in commercial projects
- Modify the code for your needs
- Distribute modified versions
- Include in your add-in packages

✅ **You SHOULD:**
- Give credit to the original author
- Include the license file with distributions
- Document any modifications you make

❌ **You CANNOT:**
- Hold the author liable for any damages
- Use the author's name for endorsement without permission

---

## লাইসেন্স

```
MIT লাইসেন্স

কপিরাইট (c) ২০২৪ আবিদ আল জামী

এই সফটওয়্যার এবং সংশ্লিষ্ট ডকুমেন্টেশন ফাইলগুলির ("সফটওয়্যার") একটি কপি প্রাপ্ত যেকোনো ব্যক্তিকে
বিনামূল্যে এখানে অনুমতি দেওয়া হচ্ছে, ব্যবহার, অনুলিপি, পরিবর্তন, একত্রিত, প্রকাশ, বিতরণ, সাবলাইসেন্স
এবং/অথবা সফটওয়্যারের কপি বিক্রয় সহ সীমাবদ্ধতা ছাড়াই সফটওয়্যারে লেনদেন করার জন্য।

উপরের কপিরাইট নোটিস এবং এই অনুমতি নোটিস সফটওয়্যারের সমস্ত কপি বা উল্লেখযোগ্য অংশে অন্তর্ভুক্ত
থাকতে হবে।

সফটওয়্যারটি "যেমন আছে" ভিত্তিতে প্রদান করা হয়, কোনো ধরনের ওয়ারেন্টি ছাড়াই।
```

### ব্যবহারের অধিকার:

✅ **আপনি পারেন:**
- ব্যক্তিগত প্রজেক্টে এই ইঞ্জিন ব্যবহার করতে
- বাণিজ্যিক প্রজেক্টে এই ইঞ্জিন ব্যবহার করতে
- আপনার প্রয়োজন অনুযায়ী কোড পরিবর্তন করতে
- পরিবর্তিত সংস্করণ বিতরণ করতে
- আপনার অ্যাড-ইন প্যাকেজে অন্তর্ভুক্ত করতে

✅ **আপনার উচিত:**
- মূল লেখককে কৃতিত্ব দিন
- বিতরণের সাথে লাইসেন্স ফাইল অন্তর্ভুক্ত করুন
- আপনার করা যেকোনো পরিবর্তন ডকুমেন্ট করুন

❌ **আপনি পারবেন না:**
- কোনো ক্ষতির জন্য লেখককে দায়ী করতে
- অনুমতি ছাড়া লেখকের নাম সমর্থনের জন্য ব্যবহার করতে

---

## Credits and Acknowledgments

### Author:
**Abid Al Jamee**
- VBA Engine Developer
- Bangladesh Bengali Calendar Implementation

### Special Thanks:
- Government of Bangladesh for official Bengali calendar specifications
- Microsoft VBA team for robust calendar APIs
- Open source community for UTF-8 file handling techniques

---

## Contact and Support

### Get Help:
- **Documentation**: Read this file thoroughly
- **Issues**: Check the Troubleshooting section
- **Questions**: Create detailed issue reports with:
  - Windows version
  - Word version
  - Exact error message
  - Input date that caused error
  - Resource file status

### Feature Requests:
Suggestions for improvements are welcome! Consider requesting:
- Additional calendar systems (Gregorian Julian, Persian, etc.)
- More language scripts (Devanagari, Thai, etc.)
- Date range validation
- Bulk date conversion utilities
- GUI interface for date conversion

---

## Appendix A: Complete Example Macro

```vba
Sub CompleteExample_DateLetterhead()
    '=========================================
    ' COMPLETE EXAMPLE: Multi-Calendar Letterhead
    '=========================================
    
    ' Input date
    Dim inputDate As String
    inputDate = "15/04/2024"  ' পহেলা বৈশাখ ১৪৩১
    
    ' Convert date
    Call ABID__DateENGINE(inputDate)
    
    ' Build letterhead content
    Dim content As String
    content = "অফিসিয়াল পত্র" & vbCrLf & vbCrLf
    
    ' Date in three calendars
    content = content & "তারিখ:" & vbCrLf
    content = content & "বাংলা: " & Abid_DEngine_BanglaDate_dayNum_bn_pre & " " & _
              Abid_DEngine_BanglaDate_monthName_bn & " " & _
              Abid_DEngine_BanglaDate_yearNum_ar & " " & _
              Abid_DEngine_BanglaDate_Era_bn & vbCrLf
              
    content = content & "English: " & Abid_DEngine_EnglishDate_monthName_en & " " & _
              Abid_DEngine_EnglishDate_dayNum_bn & ", " & _
              Abid_DEngine_EnglishDate_yearNum_ar & " " & _
              Abid_DEngine_EnglishDate_Era_en & vbCrLf
              
    content = content & "হিজরি: " & Abid_DEngine_ArabicDate_dayNum_bn & " " & _
              Abid_DEngine_ArabicDate_monthName_bn & " " & _
              Abid_DEngine_ArabicDate_yearNum_ar & " " & _
              Abid_DEngine_ArabicDate_Era_bn & vbCrLf
    
    content = content & "বার: " & Abid_DEngine_BanglaDate_weekName_bn & vbCrLf & vbCrLf
    
    content = content & "বিষয়: [আপনার বিষয় এখানে লিখুন]" & vbCrLf & vbCrLf
    
    content = content & "প্রিয় মহোদয়," & vbCrLf & vbCrLf
    
    ' Insert into document
    Selection.TypeText content
    
    MsgBox "Letterhead created successfully!", vbInformation
End Sub
```

---

## Appendix B: Resource File Template

Save this as `ABID_DateENGINE_resources.txt` with UTF-8 encoding:

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
মাঘ|Magh|مাغ
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

---

## Quick Reference Card

### Main Function Call:
```vba
Call ABID__DateENGINE("DD/MM/YYYY")
```

### Most Common Variables:

**Bengali Date:**
- `Abid_DEngine_BanglaDate_dayNum_bn_pre` → ১লা, ৫ই, ১৯শে
- `Abid_DEngine_BanglaDate_monthName_bn` → বৈশাখ
- `Abid_DEngine_BanglaDate_yearNum_ar` → ١٤٣١
- `Abid_DEngine_BanglaDate_weekName_bn` → শনিবার

**English Date:**
- `Abid_DEngine_EnglishDate_dayNum_bn` → 15
- `Abid_DEngine_EnglishDate_monthName_en` → April
- `Abid_DEngine_EnglishDate_yearNum_ar` → ٢٠٢٤

**Arabic Date:**
- `Abid_DEngine_ArabicDate_dayNum_bn` → 08
- `Abid_DEngine_ArabicDate_monthName_ar` → شوال
- `Abid_DEngine_ArabicDate_yearNum_ar` → ١٤٤٥

### Standard Formats:

**Full Bengali:** `dayNum_bn_pre + monthName_bn + yearNum_ar + Era_bn`
**Full English:** `monthName_en + dayNum_bn + yearNum_ar + Era_en`
**Full Arabic:** `dayNum_bn + monthName_ar + yearNum_ar + Era_ar`

---

**End of Documentation**

**ডকুমেন্টেশন সমাপ্ত**

---

*Last Updated: April 2024*  
*Version: 2.0*  
*Author: [Abid Al Jamee](https://facebook.com/abidaljamee)*  
*License: MIT*
