# DentBridge

Zubtexnik laboratoriyasi uchun platforma: zakazchik (shifokor, klinika), laboratoriyaning barcha xodimlari (kasbi bo'yicha), kassir va rahbarni bitta joyda bog'laydi. Telegram Mini App dizayni asosidagi web versiya.

- `index.html` — bosh sahifa
- `app.html` — web ilova (ma'lumotlar brauzerda, `localStorage` → `dentbridge_v3`)
- `kassa-import.js` — oylik kassa Excel fayllarini bazaga o'tkazuvchi modul
- `dental-kopruk.html` — dental ko'prik haqida ma'lumot sahifasi

## Rollar

| Rol | Nima qiladi |
|---|---|
| Zakazchik | Naryad beradi, holatini kuzatadi, tayyor ishni qabul qiladi yoki tuzatishga qaytaradi, qarzini ko'radi |
| Laboratoriya xodimi | O'z kasbidagi bosqichlarni oladi va topshiradi, oylik ishbay daromadini ko'radi |
| Kassir | Kunlik kassa: kirim (naqd/plastik), avans/oylik, xarajat, bank, boshqa kirimlar |
| Rahbar | Hammasi + xodimlar, arizalar, prays-list, hisobotlar, Excel import, zaxira nusxa |

Xodim kasblari: gipschi, CAD modeler (exocad), CAM operator / frezerchi, karkas texnigi, quyuvchi, keramist, implant texnigi, olinadigan protez texnigi, ortodontik texnik, pardozchi, kuryer.

## Baza tuzilmasi

- `people` — barcha odamlar: `kind` = `zak` | `tex` | `kas` | `adm`, xodimda `prof` (kasb), zakazchida `disc` (chegirma) va `pct`/`pp` (naqd/plastik ulushi, Excel'dagi «Plan»)
- `price` — ish turlari: narx, birlik va bosqichlar ro'yxati `[kasb, ishbay haqi]`
- `orders` — naryadlar: zakazchi, bemor, tishlar, birlik, rang, muddat, summa, bosqichlar (`kasb`, `xodim`, `holat`, `sana`), chat
- `days` — kassa, har kun uchun: `in` [shaxs, naqd/plastik, summa, naryad], `adv` [shaxs, summa], `exp` [izoh, summa], `bank`, `oth` [izoh, naqd/plastik, summa]

Kunlik hisob Excel'dagi formulalar bilan bir xil: *Kassada naqd = naqd − avans − xarajat + boshqa naqd kirim*, *Egasiga = kassada naqd − bank*, *Olinmagan = ulush (naqd) − avans*.

## Excel import

Rahbar → Ko'proq → Baza → «Excel'dan import». Oylik `.xlsx` fayllarni tanlang (bir nechtasini birdaniga). Har bir «2026 Avgust (4)» ko'rinishidagi kunlik varaqdan topshirilgan pul, Plan ulushi, avanslar, xarajatlar, bank va boshqa kirimlar olinadi. Excel'dagi summalar ming so'mda deb hisoblanadi. Fayllar brauzerda o'qiladi va hech qayerga yuborilmaydi (Excel o'qish kutubxonasi cdnjs'dan yuklanadi).

Statik sayt: qurish kerak emas. GitHub Pages: Settings → Pages → Branch: `main` / root.
