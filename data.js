// قاعدة بيانات الدورة
const COURSE_DATA = {
  1: {
    title: 'مقدمة في Excel وشريط الأدوات (Ribbons)',
    desc: 'التعرف على الواجهة، شريط الأدوات (Ribbons)، وأهمية Excel في عالم الأعمال.',
    story: 'نبدأ رحلتنا بكسر الجدار الأول مع الإكسل. سنتعرف على الواجهة بالتفصيل، ونغوص في شريط الأدوات (Ribbons) لنشرح كل أيقونة وكيفية استخدامها للتحكم الكامل بأوراق العمل.',
    objectives: ['التعرف على واجهة Excel وأدواته', 'فهم شريط الأدوات (Ribbons) ووظيفة كل أيقونة', 'التنقل بكفاءة داخل أوراق العمل'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQQPNsV0qfZ8Sr6eZQHbHsgwAZa32zBChYmvLDUSryy-l8s?width=1920&height=1080',
    files: [
      { name: 'Taps Part 1', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQB0qdk3vnJiQpZhWiPTbpEaAZo1yC2eLg_qCffdXALYGDk?download=1', type: 'Excel Workbook' }  
    ]
  },
  2: {
    title: 'الصيغ والدوال الأساسية',
    desc: 'SUM، AVERAGE، COUNT وكيفية توظيفها في الحياة العملية.',
    story: 'في هذه المحاضرة سننتقل من مجرد إدخال البيانات إلى جعل Excel يحسب نيابة عنك. سنتعلم أهم الدوال الأساسية التي لا غنى عنها في أي بيئة عمل.',
    objectives: ['فهم الهيكل الصحيح لكتابة الصيغ', 'تطبيق دوال الجمع والمتوسط (SUM, AVERAGE)', 'استخدام دوال العد (COUNT) وتجنب الأخطاء'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQR4-XD-9DlAT5NS_MleLUCVAWu3s7mXAb2kUBrK3e4akCc?width=1920&height=1080', 
    files: [
      { name: 'Taps Part 2', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQCXteAQFYtZSp8adjzNOA8UARV4JNgWt-mcmfSqmtcesm8?download=1', type: 'Excel Workbook' },
      { name: 'كل معادلات الاكسل شرح بالعربي', url: 'https://1drv.ms/b/c/a1a27fad04d192e9/IQAjEFtIAbcwQKfQsQTUI3KDAZy4DqJ5ZRPm3wCS7H8gNl4?download=1', type: 'PDF Document' }
    ]
  },
  3: {
    title: 'مهارات توفير الوقت واللصق الخاص',
    desc: 'تعديل الشيتات المتعددة، اختصارات هامة، اللصق الخاص (Paste Special) ومقدمة للرسوم البيانية.',
    story: 'محاضرة دسمة جداً! سنتعلم خدعاً توفر الكثير من الوقت مثل التعديل على عدة أوراق عمل في نفس اللحظة (Multiple Sheets)، وأسرار اللصق الخاص، بالإضافة إلى مقدمة في الرسوم البيانية وأهم اختصارات الكيبورد.',
    objectives: ['التعديل على أوراق عمل متعددة دفعة واحدة (Ctrl + Select)', 'استخدام اللصق الخاص (Paste Special) باحترافية', 'إنشاء الرسوم البيانية (Charts) واستخدام اختصارات الإكسل'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQRo2LpEw6SgQ5cy-iPyhJhIAWsN2YTN79rxvR_KRfzoSDc?width=1920&height=1080',
    files: [
      { name: 'Editing Multiple Sheets', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQCe_r8UCHT3T5dZDE7iZbj7ASspBimLh6kukcelzqYRNaY?download=1', type: 'Excel Workbook' },
      { name: 'جميع اختصارات الاكسل', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQA7jo5VIwEHQKif0xKZT8cJAXJ61kD8GP5bKOSkKpmS7dA?download=1', type: 'Excel Workbook' },
      { name: 'Charts & Data Visualization', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQDDlpkxdaHNQr3qnuER0yg-AWPrhU4GpNkcirr1l4Mg8JM?download=1', type: 'Excel Workbook' },
      { name: 'Copy & Paste Special', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQB8i-J1kgSASJ_kXjUSeeUYAUdQ7HXJekEpaH4f-YRxZ-U?download=1', type: 'Excel Workbook' },
      { name: 'Assignment 2 (الواجب الأول)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQAv5oSuJ24MTbdJsQZLhpjpAalv2wc5Yc-3-lQOKBICfmw?download=1', type: 'Excel Workbook' }
    ]
  },
  4: {
    title: 'التنسيق الشرطي المتقدم والدوال الإحصائية',
    desc: 'احتراف التنسيق الشرطي (Conditional Formatting) ودوال العد والمجاميع المتقدمة.',
    story: 'نبدأ برفع مستوى الاحتراف. سنتعمق في التنسيق الشرطي الديناميكي لنجعل التقارير تتفاعل مع البيانات تلقائياً. وسنستكشف قوة دوال العد المتقدمة (Counts) وأدوات المجاميع (Subtotal & SumProduct) التي تفصل المحترف عن المبتدئ.',
    objectives: ['إتقان التنسيق الشرطي الديناميكي (Dynamic Design)', 'استخدام دوال Subtotal و SumProduct باحتراف', 'احتراف دوال العد المتقدمة (Mastering Counts)'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQRO-5Vwp87MQZl8bklQ2UYHAQoqzY1n4JI3XGkiX_S0SII?width=1920&height=1080',
    files: [
      { name: '08-Dynamic Design Conditional', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQAt009GwcxMS6Fy5yYcKr2YAYtWjlqVH2209NltgxX3c88?download=1', type: 'Excel Workbook' },
      { name: '10-Subtotal & SumProduct', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQAcXPk6KUb_TIPGUc9VMmJlAebIeh-immuB32J2qXApYvs?download=1', type: 'Excel Workbook' },
      { name: '11-Conditional Formatting Extras', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQBuGyejtcaETJX6q1CggYLhAV8yF3kXbQLoWKxZBFvLwqo?download=1', type: 'Excel Workbook' },
      { name: '12-On-Mastering Counts & More', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQCUA73oU3m5TbULoSHNn282AXg2khtNGi4yCoC-DKh13no?download=1', type: 'Excel Workbook' },
      { name: 'Session Exercise', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQCXq0RY3rhSQLB_0_lIhcJKAU8K_tlMbfsdcXYcbs8dMWo?download=1', type: 'Excel Workbook' }
    ]
  },
  5: {
    title: 'سيمفونية البحث والتحكم بالبيانات',
    desc: 'تسهيل المعادلات بـ Name Manager، التحكم بالمدخلات بـ Data Validation، واحتراف دالة VLOOKUP الأسطورية.',
    story: 'هل أنت مستعد لنقلة نوعية؟ في هذه المحاضرة الحماسية سنكتشف كيف نجعل حياتنا والمعادلات المعقدة أسهل بكثير باستخدام الـ Name Manager. سنتعلم أيضاً كيف نفرض سيطرتنا على مدخلات المستخدمين عبر Data Validation (ونضيف لمستنا الخاصة برسائل خطأ مضحكة!). ومسك الختام سيكون مع "سيمفونية البحث".. نعم، سنتعلم معاً أعظم دالة في الإكسل: VLOOKUP، لنجعل البيانات تبحث عن نفسها بنفسها!',
    objectives: ['تسهيل قراءة وكتابة المعادلات المعقدة باستخدام Name Manager', 'تقييد إدخال البيانات وإنشاء رسائل خطأ مخصصة عبر Data Validation', 'احتراف دالة VLOOKUP لربط البيانات والبحث عنها بلمح البصر'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQRyTqZGsYWpQ4BFaLNCB5sJARxjBruaQnd24lJXmBjV6Eg?width=1920&height=1080', 
    files: [
      { name: 'Name Manager', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQBWNL3hDXvDSo3ZWsv8N0xYAb5CXZOv9sLVSldBGIt0Dzw?download=1', type: 'Excel Workbook' },
      { name: 'Data Validation', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQD8i3FCYg7zQaDiET4Y4f9dAS3-wY3ZfGRZutmbuReT8S4?download=1', type: 'Excel Workbook' },
      { name: 'Data Lookup Symphony (VLOOKUP)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQANLuW7YOMDQLqBv-LgrwswARDyB4Kk8-AnhDWFyxn-tx0?download=1', type: 'Excel Workbook' }
    ]
  },
  6: {
    title: 'خفايا الطباعة، التنقل السحري، والاحتراف المطلق لـ VLOOKUP',
    desc: 'اتقان الطباعة الاحترافية، التنقل السريع باستخدام Go To Special، وخفايا البحث المتقدم.',
    story: 'استعد لمحاضرة مليئة بالخدع السحرية! هل تعاني من طباعة الجداول وتنسيقها؟ سنقضي على هذه المشكلة تماماً ونجعلك تطبع أي ملف باحترافية تامة مع ضبط الرأس والتذييل. سنتعلم أيضاً كيف "نطير" داخل الإكسل ونصل لأي معلومة مخفية بلمح البصر باستخدام "Go To Special". ولأن VLOOKUP هي العشق، سنعود إليها لنعزف معاً ألحاناً متقدمة ونكشف أسرارها العميقة!',
    objectives: ['احتراف طباعة الملفات وتنسيق الرأس والتذييل (Header & Footer)', 'التنقل السريع والذكي واستكشاف خبايا أداة Go To Special', 'الغوص في الحالات المتقدمة والاحترافية لدالة VLOOKUP'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQRi-N0K7oBRQKvK1gQCuB14AdogNfQ2FV5znFGsJCpaazM?width=1920&height=1080', 
    files: [
      { name: 'Print Perfection (الطباعة الاحترافية)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQAFmROmkaRDS7Ti-RqMKTBVAYNgZwUOjXbqXPFUK4SqQH0?download=1', type: 'Excel Workbook' },
      { name: 'Excel Navigation (التنقل الذكي)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQCsWVz8cAQAQ6jraEhAC5W3ARGwRDpzK1bf1I8ggPjQ2js?download=1', type: 'Excel Workbook' },
      { name: 'Data Lookup Symphony (VLOOKUP المتقدم)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQBii14BNy8TS7UOk05UdahLAV0nBcGAiuJ-VRkMKYGVP2U?download=1', type: 'Excel Workbook' },
      { name: 'Assignment 3 (الواجب الثالث)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQCG7rrCA_ojRJ8KaAGz9xY0AVzn0klBEt4figKZ3FTfkh0?download=1', type: 'Excel Workbook' }
    ]
  },
  7: {
    title: 'المصفوفات الديناميكية ودالة XLOOKUP الخارقة',
    desc: 'رحلة دمج معادلات FILTER و SORT و UNIQUE للوصول للجيل الجديد من البحث بـ XLOOKUP.',
    story: 'استعدوا للصدمة الإيجابية! في هذه المحاضرة سنعيش رحلة جنونية مع المصفوفات الديناميكية؛ سندمج دوال مثل FILTER و UNIQUE و SORT و VSTACK معاً لتشكيل خلطة سحرية تستخرج أي نتيجة مهما كانت معقدة. وبعدها، سنودع الطرق التقليدية ونستقبل الملك المتوج "XLOOKUP"، لنفصص كل جزء فيه ونكتشف الفروقات الجوهرية التي تجعله البديل الأقوى والأكثر مرونة مقارنة بـ VLOOKUP. هذه المحاضرة ستغير طريقة عملك للأبد!',
    objectives: ['استخراج البيانات وتصفيتها باحترافية باستخدام دالة FILTER', 'دمج المصفوفات الديناميكية (UNIQUE, SORT, CHOOSECOLS, VSTACK) في معادلة واحدة جبارة', 'فهم الفروقات الجوهرية واحتراف جميع خصائص دالة XLOOKUP السحرية'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQS8aDI7l06gQJbUzXiJMZtVAUheJApM9YYE-5aRmK9vsoA?width=1920&height=1080', 
    files: [
      { name: 'Using Filter (رحلة المصفوفات)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQBWtYicqDGjTYk1E_eolA6jARzZZMMWm8rhlk9K5EXiaXc?download=1', type: 'Excel Workbook' },
      { name: 'XLOOKUP (الملك المتوج)', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQD7OYzI-MDeSqXK7O0kw6XzARcLMv007Jwg7pGJasq8M30?download=1', type: 'Excel Workbook' }
    ]
  },
  8: {
    title: 'سحر Flash Fill، الشروط المتداخلة، ودوال النصوص و INDEX MATCH',
    desc: 'تسريع العمل بـ Flash Fill، إتقان دالة IF المتداخلة، التلاعب بالنصوص، والسر الأعظم INDEX و MATCH.',
    story: 'جاهزون لجرعة مكثفة من أسرار المحترفين؟ سنبدأ بسحر خاصية Flash Fill التي ستنجز ساعات من العمل في ثانية واحدة وبطريقة مرتبة جداً! ثم سنبني منطقاً حديدياً بتداخل معادلات IF المتعددة لنعالج أعقد الحالات العملية. ولأن البيانات غالباً ما تأتي فوضوية، سنسيطر عليها تماماً باستخدام ترسانة من دوال النصوص القوية (مثل TEXTSPLIT و TRIM وغيرها). وأخيراً، سنكشف عن السلاح السري للمحللين المتقدمين: دمج معادلة INDEX مع MATCH للبحث بمرونة لا حدود لها!',
    objectives: ['تسريع المهام الروتينية وتنظيف البيانات باستخدام الذكاء المدمج لخاصية Flash Fill', 'بناء شروط منطقية معقدة ومتداخلة باستخدام IF Statement لسيناريوهات العمل المختلفة', 'معالجة وتنظيف النصوص المعقدة باستخدام ترسانة الدوال (TRUNC, TRIM, TEXTSPLIT, etc.)', 'احتراف البحث المتقدم والبديل الأقوى لـ VLOOKUP باستخدام دمج INDEX و MATCH'],
    videoUrl: 'https://1drv.ms/v/c/a1a27fad04d192e9/IQQJTmLRfnJBSY0Sp9WX8A4YAe8u0_S9RA301a28r9177TM?width=1920&height=1080', 
    files: [
      { name: '20-Flash Fill', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQB9UN6ggTfbQ7YJ1_S7DrTwAUUd1knuMK3hJZvpH0I5R4k?download=1', type: 'Excel Workbook' },
      { name: 'If statement and IF combined', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQDdtuy7vk2aSJHJZ5U5sASTATCQW6Mn8xvkg_wG4a0OSBQ?download=1', type: 'Excel Workbook' },
      { name: '22-Text formulas', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQBzhLqNBDQ3QJ_lOx7HD4ZbAc45Mq50w5wI_B8tsmQYQng?download=1', type: 'Excel Workbook' },
      { name: 'Index Match Magic', url: 'https://1drv.ms/x/c/a1a27fad04d192e9/IQClppGCEXrET4LQeX8po57YAYDwipzamtazNJQikYjZJXg?download=1', type: 'Excel Workbook' }
    ]
  },
  9: {
    title: 'Power Query والأتمتة',
    desc: 'استيراد وتحويل البيانات بقوة لا تصدق وأتمتة المهام.',
    story: 'الأداة السحرية والمخفية في Excel! سنتعلم كيف ننظف ونرتب البيانات الفوضوية القادمة من مصادر مختلفة بشكل آلي، وكيف نوفر ساعات من العمل الروتيني.',
    objectives: ['استيراد البيانات من مصادر متعددة وتنظيفها', 'دمج ملفات مختلفة في جدول واحد', 'مقدمة سريعة في أتمتة المهام المتكررة'],
    videoUrl: '', 
    files: []
  },
  10: {
    title: 'مشروع نهائي متكامل',
    desc: 'تطبيق شامل على كل ما تعلمته — إنشاء لوحة تحكم (Dashboard).',
    story: 'المحطة الأخيرة في رحلتنا! سنقوم بدمج كل المهارات التي تعلمناها في الدورة لبناء لوحة تحكم تفاعلية واحترافية من الصفر وحتى الإخراج النهائي.',
    objectives: ['التخطيط الهيكلي للوحة التحكم', 'ربط الدوال والجداول المحورية ببعضها', 'التصميم النهائي والإخراج البصري الاحترافي'],
    videoUrl: '', 
    files: []
  }
};
