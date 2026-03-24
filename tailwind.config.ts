import type { Config } from "tailwindcss";
export default { content:["./app/**/*.{ts,tsx}","./components/**/*.{ts,tsx}","./lib/**/*.{ts,tsx}"], theme:{extend:{colors:{brand:{orange:"#ff7a17",navy:"#0b1736",line:"#d8dde8"}}}}, plugins:[] } satisfies Config;
