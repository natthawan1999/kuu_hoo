# KUUHOO Demo

ระบบนับสต็อก · Demo version

## วิธี Deploy บน Vercel ผ่าน GitHub

### 1. สร้าง GitHub repo
```bash
git init
git add .
git commit -m "init: kuuhoo demo"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/kuuhoo-demo.git
git push -u origin main
```

### 2. Deploy บน Vercel
1. ไปที่ [vercel.com](https://vercel.com) → **Add New Project**
2. Import repo `kuuhoo-demo` จาก GitHub
3. Vercel จะ detect เป็น **Vite** อัตโนมัติ
4. กด **Deploy** — เสร็จใน ~1 นาที

### Settings ที่ Vercel (ปล่อย default ได้เลย)
| | |
|---|---|
| Framework | Vite |
| Build Command | `npm run build` |
| Output Directory | `dist` |
| Install Command | `npm install` |

## Run local
```bash
npm install
npm run dev
```
