# بوابة القسم الداخلية — Jazan Department Portal

هذا المستودع يحتوي على **الكود المصدري فقط** لبوابة القسم الداخلية وجميع المنصات المرتبطة بها (Google Apps Script).

⚠️ **لا يحتوي هذا المستودع على أي بيانات فعلية** (لا شيتات بيانات، لا ملفات شواهد/مخرجات، لا معلومات شخصية كأرقام هوية أو جوالات أو أسماء موظفين). فقط أكواد `Code.gs` و `Index.html` الخاصة بكل منصة.

## البوابة (Portal)

ملف [`index.html`](./index.html) في جذر المستودع هو بوابة القسم — واجهة موحّدة (قائمة جانبية + iframe) تجمع روابط جميع المنصات مصنّفة حسب الفئة، وتُستضاف على GitHub Pages:

**https://jazan-edu.github.io/jazan-department-portal/**

لإضافة منصة جديدة: أضف عنصراً إلى مصفوفة `PLATFORMS` في `index.html` (الفئة، الاسم، الرابط) وأيقونته في `ICONS` بنفس الترتيب. المنصة بدون رابط (`url: null`) تظهر تلقائياً كـ "قريباً".

## المنصات (Platforms)

| المجلد | المنصة | الفئة |
|---|---|---|
| [`attendance-musanid`](./platforms/attendance-musanid) | منصة مساند حضوري (نظام إدارة الغياب والمساءلات) | الحضور والمساءلات |
| [`performance-evidence`](./platforms/performance-evidence) | منصة شواهد تنفيذ الأهداف للمشرفين ورؤساء الوحدات | الإشراف التربوي |
| — | منصة مساند مدرستي | الإشراف التربوي |
| [`performance-cycle-1447`](./platforms/performance-cycle-1447) | منصة دورة الأداء الوظيفي 1447 | الأداء الوظيفي والتميز |
| [`self-evaluation-tracking`](./platforms/self-evaluation-tracking) | منصة متابعة حالة التقويم الذاتي | الأداء الوظيفي والتميز |
| [`excellence-modifications`](./platforms/excellence-modifications) | منصة استقبال التعديلات على منصة التميز | الأداء الوظيفي والتميز |
| [`chemical-safety`](./platforms/chemical-safety) | منصة حصر المواد الكيميائية وقواعد السلامة المرورية | النماذج والسلامة |
| — | منصة طلبات التحول عن بعد | طلبات المدارس |

## ملاحظة تقنية مهمة

جميع منصات `doGet()` تحتاج السطر التالي حتى تُفتح داخل iframe (كما في البوابة):

```javascript
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
```

وبعد أي تعديل على الكود في محرر Apps Script، لازم إنشاء **إصدار جديد (New version)** من قائمة Deploy → Manage deployments حتى ينعكس التعديل على الرابط المنشور فعلياً.
