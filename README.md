# Workforce Suite — ដំណើរការនៅលើកុំព្យូទ័រខ្លួនឯង

## តម្រូវការ
- ដំឡើង [Node.js](https://nodejs.org) (កំណែ 18 ឡើងទៅ) ជាមុនសិន

## ជំហានដំណើរការ

1. ស្រង់ (unzip) ឯកសារនេះ រួចបើកកម្មវិធី VS Code ត្រង់ folder នេះ
2. បើក Terminal ក្នុង VS Code (Terminal → New Terminal) រួចវាយ៖

   ```bash
   npm install
   ```

   (ធ្វើតែម្តងគត់ ដើម្បីទាញយក library ដែលត្រូវការ)

3. បន្ទាប់មកវាយ៖

   ```bash
   npm run dev
   ```

4. Terminal នឹងបង្ហាញ link ដូចជា `http://localhost:5173` — ចុច Ctrl+Click (ឬចម្លងទៅ Browser) ដើម្បីមើលកម្មវិធី

## ចំណាំសំខាន់ៗ

- **ទិន្នន័យរក្សាទុកនៅឯណា?** កូដដើមប្រើ `window.storage` ដែលដំណើរការតែក្នុង Claude.ai ប៉ុណ្ណោះ។ ខ្ញុំបានបន្ថែមឯកសារ `src/storage.js` ដែលក្លែងធ្វើមុខងារដូចគ្នា ប៉ុន្តែរក្សាទុកទៅ **localStorage របស់ Browser** វិញ (ន័យថា ទិន្នន័យស្ថិតនៅតែលើ Browser/ម៉ាស៊ីននោះ មិនបានចែករំលែកទៅអ្នកប្រើផ្សេងទេ)។
- ប្រសិនបើប្រូចង់ភ្ជាប់ទៅ database ពិតប្រាកដ (ដូចជា Firebase, Supabase, ឬ backend ខ្លួនឯង) ក្រោយពេលនេះ គ្រាន់តែកែឯកសារ `src/storage.js` ប៉ុណ្ណោះ គ្មានចាំបាច់កែ `App.jsx` ទេ។
- ដើម្បីចេញជាឯកសារសម្រាប់ដាក់លើ server ពិត (production) ប្រើ៖

  ```bash
  npm run build
  ```

  លទ្ធផលនឹងចេញនៅ folder `dist/` ដែលអាចយកទៅ upload លើ hosting ណាមួយ (Netlify, Vercel, ។ល។)

## គណនីសាកល្បង
- Admin — ពាក្យសម្ងាត់: `admin123`
- បុគ្គលិក — EMP-001/1001, EMP-002/1002, EMP-003/1003
