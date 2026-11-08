# نسخة سطح المكتب (ويندوز)

برنامج ويندوز يفتح المنظومة في نافذة مستقلة باسمها وأيقونتها، دون شريط المتصفح.

- **التحميل:** [BregaSales-Setup.exe](https://github.com/melmzughi-svg/Al-brega-Sells-Company/releases/latest/download/BregaSales-Setup.exe) (مثبّت) أو [BregaSales-Portable.exe](https://github.com/melmzughi-svg/Al-brega-Sells-Company/releases/latest/download/BregaSales-Portable.exe) (محمول بلا تثبيت).
- **التحديث:** البرنامج يفتح المنظومة الحية على GitHub Pages، فكل تحديث للموقع يظهر فيه تلقائياً ولا حاجة لإعادة التثبيت. يُعاد بناء المثبّت فقط عند تغيير ملفات هذا المجلد.
- **دون اتصال:** يحتاج أول تشغيل إلى الإنترنت، ثم يعمل دون اتصال عبر نفس آلية الموقع (Service Worker).
- **تنبيه ويندوز:** المثبّت غير موقّع رقمياً، فقد يظهر «Windows protected your PC». اضغط «More info» ثم «Run anyway».

## للمطوّر

```bash
cd desktop
npm install
npm start                                   # يشغّل البرنامج على المنظومة الحية
BREGA_APP_URL=http://127.0.0.1:8765/ npm start   # على نسخة محلية (يعمل فقط من المصدر، لا في البرنامج المثبّت)
npm run dist:win                            # يبني المثبّت في desktop/dist (يفضّل على ويندوز)
```

البناء والنشر تلقائيان عبر `.github/workflows/desktop.yml`: طلبات الدمج تبني فقط، والدمج في `main` ينشر إصداراً باسم `desktop-v<الإصدار>`. لإصدار مثبّت جديد ارفع رقم `version` في `desktop/package.json`.
