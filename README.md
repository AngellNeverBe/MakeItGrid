# MakeItGrid

20×20 Chinese manuscript paper (&#26041;&#26684;&#32440;) preview generator — Tauri v2 desktop app.

Convert `.docx` files into traditional 20×20 grid manuscript paper HTML previews.

---

&#23558; `.docx` &#25991;&#26723;&#36716;&#25442;&#20026; 20×20 &#26631;&#20934;&#26041;&#26684;&#32440; HTML &#39044;&#35272;&#12290;

## Features / &#21151;&#33021;

- **Import** `.docx` files via drag & drop or native file dialog
- **Generate** grid manuscript previews with one click
- **View** previews in-app — no external browser needed
- **Frameless** window with custom titlebar controls
- Supports Chinese punctuation pairing, digit/letter grouping, and proper grid alignment

---

- **&#23548;&#20837;** &#8212; &#25302;&#25341;&#25110;&#21407;&#29983;&#25991;&#20214;&#23545;&#35805;&#26694;&#23548;&#20837; .docx
- **&#19968;&#38190;&#29983;&#25104;** &#8212; &#28857;&#20987;&#29983;&#25104;&#26041;&#26684;&#32440;&#39044;&#35272;
- **&#24212;&#29992;&#20869;&#39044;&#35272;** &#8212; &#26080;&#38656;&#22806;&#37096;&#27983;&#35272;&#22120;
- **&#26080;&#36793;&#26694;&#31383;&#21475;** &#8212; &#33258;&#23450;&#20041;&#26631;&#39064;&#26639;&#25511;&#20214;

## Build / &#26500;&#24314;

Requires Rust (GNU toolchain), Node.js, Python 3.

```bash
npm install
npx tauri build
```

## Tech Stack / &#25216;&#26415;&#26632;

Tauri v2 + Rust + Vanilla JS + Python (python-docx)

## License

MIT
