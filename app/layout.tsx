import "./globals.css";
export const metadata = { title:"ASN TOOL GM", description:"Clean rebuild", icons:{icon:"/icon.png", apple:"/icon.png"} };
export default function RootLayout({children}:{children:React.ReactNode}){ return <html lang="en"><body>{children}</body></html>; }
