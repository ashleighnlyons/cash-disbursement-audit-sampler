import type { Metadata } from "next";
import "./globals.css";
export const metadata:Metadata={title:"Cash Disbursement Audit Sampler",description:"Generate reproducible cash disbursement audit selections from paid-expense reports."};
export default function RootLayout({children}:Readonly<{children:React.ReactNode}>){return <html lang="en"><body>{children}</body></html>}
