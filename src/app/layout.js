import { FinancialProvider } from '@/context/FinancialContext';
import './globals.css';
export default function RootLayout({ children }) {
    return (
        <html lang="en">
            <head>
                <title>Fina AI — Multi-Template Financial Platform</title>
            </head>
            <body suppressHydrationWarning className="antialiased min-h-screen">
                <FinancialProvider>
                    {children}
                </FinancialProvider>
            </body>
        </html>
    );
}
