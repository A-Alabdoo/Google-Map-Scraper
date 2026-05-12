# أداة استخراج بيانات الأنشطة التجارية من Google Maps

سكربت Python لجمع بيانات الأنشطة التجارية (الاسم، الهاتف، البريد الإلكتروني، الموقع الإلكتروني، رابط الخريطة) من خرائط Google بناءً على قائمة كلمات مفتاحية، وتصدير النتائج إلى ملفات Excel.

---

## 🧱 صعوبة عمل Scraping من Google Maps بالطرق التقليدية

استخراج البيانات من Google Maps يُعدّ من أصعب أنواع الـ Web Scraping للأسباب التالية:

1. **الصفحة ديناميكية بالكامل (JavaScript-heavy):** خرائط Google لا تُرجع HTML جاهز، بل يتم تحميل البيانات عن طريق طلبات XHR وعرضها عبر React/JS. لذلك أدوات مثل `requests` + `BeautifulSoup` لا تعمل إطلاقًا، لأنها تستلم صفحة فارغة.
2. **التمرير اللانهائي (Infinite Scroll):** نتائج البحث لا تظهر دفعة واحدة، بل يجب تمرير اللوحة الجانبية (`div[role="feed"]`) مرارًا حتى تظهر بقية الأماكن، وكل تمرير يُطلق طلب شبكي جديد.
3. **شاشة الموافقة (Consent Screen):** Google تعيد التوجيه إلى `consent.google.com` خصوصًا في أوروبا، ولا يمكن المتابعة قبل الضغط على "Accept all" بلغات مختلفة.
4. **أنظمة كشف البوتات المتقدمة:** Google تستخدم بصمات (Fingerprints) للمتصفح، وتفحص: User-Agent، WebGL، Canvas، Navigator properties، الإضافات، حركة الفأرة، توقيت الكتابة... إلخ. أي محاولة بـ Selenium/Playwright العاديين تُكشف خلال ثوانٍ ويظهر CAPTCHA.
5. **الـ Selectors تتغير باستمرار:** أسماء الكلاسات مثل `hfpxzc` و `DUwDvf` تتغيّر دوريًا، لذا يجب اعتماد أكثر من Selector بديل.
6. **الحظر بالـ IP والجلسة:** بعد عدد قليل من الطلبات السريعة، تبدأ Google بإرجاع صفحات فارغة أو حظر الـ IP.
7. **استخراج الإيميل غير متوفر مباشرة:** Google لا تعرض البريد الإلكتروني، فيجب الدخول إلى موقع النشاط نفسه ثم البحث في صفحات Contact / About.

---

## ✅ كيف تم حل هذه المشاكل عبر مكتبة Scrapling

مكتبة [Scrapling](https://github.com/D4Vinci/Scrapling) قدّمت حلولاً عملية لمعظم هذه التحديات:

- **`StealthySession`:** جلسة متصفح مبنية فوق Playwright لكن مع تعديلات تخفي بصمات الأتمتة (إخفاء `navigator.webdriver`، تزييف WebGL/Canvas، إعدادات اللغة والمنطقة الزمنية... إلخ).
- **`real_chrome=True`:** يستخدم متصفح Chrome الحقيقي المثبّت على الجهاز بدل Chromium، مما يجعل البصمة مطابقة 100% لمستخدم حقيقي.
- **`google_search=True`:** يفعّل وضعًا خاصًا لمحاكاة قدوم المستخدم من بحث Google، وهو أمر تتحقق منه خرائط Google.
- **`network_idle=True` و `wait_selector`:** انتظار ذكي حتى تنتهي الطلبات الديناميكية ويظهر العنصر المطلوب فعليًا، بدلًا من `time.sleep` العشوائي.
- **`page_action`:** تمرير دالة Playwright مخصصة لتنفيذ Scroll تلقائي داخل لوحة النتائج وجمع الروابط عبر `page.evaluate(JS)` مباشرة من DOM.
- **`FetcherSession(impersonate="chrome")`:** جلسة HTTP خفيفة (بديلة لـ requests) تُحاكي بصمة TLS/HTTP2 لمتصفح Chrome، تُستخدم لجلب صفحات المواقع الخارجية بسرعة دون فتح متصفح.
- **محرك Selectors قوي:** يدعم CSS بصيغة `::text` و `::attr(href)` مباشرة، ويُسهّل التعامل مع HTML الديناميكي.

---

## 🛡️ كيف يتجاوز Scrapling أنظمة كشف البوتات (Anti-Bot)

1. **بصمة متصفح حقيقية:** عبر `real_chrome=True` و `StealthySession`، يصبح المتصفح غير مميز عن متصفح المستخدم العادي على مستوى JS Fingerprinting.
2. **إخفاء علامات الأتمتة:** يقوم تلقائيًا بحذف خصائص مثل `navigator.webdriver`، وتعديل `navigator.plugins` و `navigator.languages` و `chrome.runtime` لتبدو طبيعية.
3. **محاكاة TLS/HTTP بصمة كروم (Impersonation):** `FetcherSession(impersonate="chrome")` يبني الـ TLS Handshake بنفس ترتيب Cipher Suites الذي يستخدمه Chrome، فلا يُكشف عبر JA3.
4. **انتظار شبكي ذكي بدل Sleep ثابت:** `network_idle` و `wait_selector` يجعلان السلوك يبدو طبيعيًا (تحميل الصفحة كاملة قبل التفاعل) ويتجنبان الأنماط الميكانيكية التي ترصدها أنظمة الحماية.
5. **التعامل مع شاشة الموافقة تلقائيًا:** الكود يكتشف صفحة `consent.google.com` ويضغط على "Accept all" بعدة لغات، ثم يُعيد توجيه الجلسة بشكل سلس كأن مستخدمًا حقيقيًا قد فعل ذلك.

---

## 📦 المكتبات المطلوبة وطريقة التثبيت

السكربت يعتمد فقط على **Scrapling** + متطلباتها، أما باقي الوظائف (Excel, Regex, Zip) فمستخدمة من المكتبة القياسية لـ Python.

### المتطلبات

- Python 3.10 أو أحدث.
- متصفح Google Chrome مثبّت على الجهاز (لأن `REAL_CHROME = True`).

### خطوات التثبيت

```powershell
# 1) تثبيت Scrapling مع وحدة الـ fetchers (المتصفح + HTTP)
python -m pip install "scrapling[fetchers]"

# 2) تثبيت متصفحات Playwright وأدوات Stealth التي يحتاجها Scrapling
scrapling install
```

> 💡 إذا ظهرت رسالة `ModuleNotFoundError`، السكربت سيُرشدك تلقائيًا إلى نفس الأمرين أعلاه.

---

## 🚀 دليل استخدام السكربت

### 1) تجهيز ملف الكلمات المفتاحية

أنشئ ملف `keywords.txt` في نفس مجلد السكربت، وضع كل كلمة مفتاحية في سطر مستقل، مثلًا:

```
مطاعم في دمشق
فنادق في حلب
شركات سياحة في اللاذقية
```

### 2) ضبط الإعدادات (اختياري)

في أعلى ملف [main.py](main.py#L18-L23) يمكنك تعديل:

| المتغير | الوصف | القيمة الافتراضية |
|---|---|---|
| `HEADLESS` | تشغيل المتصفح بدون نافذة مرئية | `False` |
| `REAL_CHROME` | استخدام Chrome المثبّت على الجهاز | `True` |
| `MAX_RESULTS_PER_KEYWORD` | الحد الأقصى للنتائج لكل كلمة | `700` |
| `SCROLL_PAUSE_MS` | المهلة بين كل تمرير وآخر (ميلي ثانية) | `2000` |
| `NO_NEW_RESULTS_LIMIT` | عدد محاولات التمرير قبل الإيقاف عند عدم ظهور نتائج جديدة | `8` |
| `REQUEST_TIMEOUT` | المهلة القصوى لكل طلب (ميلي ثانية) | `60000` |

### 3) تشغيل السكربت

```powershell
python main.py
```

سترى المتصفح يفتح ويبدأ بالبحث والتمرير تلقائيًا، مع طباعة تقدم العمل في الـ Terminal:

```
[+] البحث عن: مطاعم في دمشق
[+] تم تجميع 312 رابط مكان من أصل المطلوب 700
    [1/312] https://www.google.com/maps/place/...
        -> اسم المطعم | +963 11 1234567 | info@example.com
    ...
```

---

## 📊 شكل المخرجات

بعد انتهاء التشغيل، سيتم إنتاج الملفات التالية:

1. **ملف Excel لكل كلمة مفتاحية**: مثلًا `مطاعم_في_دمشق.xlsx`
2. **ملف Excel موحّد لكل النتائج**: `all_results.xlsx`

كل ملف يحتوي على الأعمدة التالية:

| Keyword | Company Name | Email | Phone | Website | Google Maps URL | Registered Email Found | Email Status |
|---|---|---|---|---|---|---|---|
| مطاعم في دمشق | مطعم النخيل | info@alnakheel.sy | +963 11 2345678 | https://alnakheel.sy | https://www.google.com/maps/place/... | Yes | Public email found on website |
| مطاعم في دمشق | مطعم الشام | Info@alsham.sy \| Sales@alsham.sy | +963 11 9876543 | https://alsham.sy | https://www.google.com/maps/place/... | No | Generated from domain because no public email was found |
| مطاعم في دمشق | مطعم بدون موقع | N/A | +963 11 1112223 | N/A | https://www.google.com/maps/place/... | No | No email and no website |

### معاني الأعمدة الخاصة

- **Registered Email Found**:
  - `Yes` → تم العثور على بريد رسمي منشور على الموقع.
  - `No` → لم يُعثر على بريد منشور (سواء تم توليده من الدومين أو لم يوجد موقع).
- **Email Status**: حالة استخراج البريد:
  - `Public email found on website` — بريد حقيقي موجود في الموقع.
  - `Generated from domain because no public email was found` — بريد مُقترَح (Info@/Sales@) مبني على الدومين.
  - `No email and no website` — لا يوجد موقع للنشاط أصلًا.

> ✅ النتائج في ملف `all_results.xlsx` تكون **بدون تكرار** (يتم إزالة المكرر بناءً على الاسم + الهاتف + الموقع + رابط الخريطة).
