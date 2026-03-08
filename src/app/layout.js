import { FinancialProvider } from '@/context/FinancialContext';
import './globals.css';
export default function RootLayout({ children }) {
    return (
        <html lang="en">
            <head>
                <title>Docty-Healthcare Smart Platform</title>
            </head>
            <body className="antialiased min-h-screen">
                <FinancialProvider>
                    {children}
                </FinancialProvider>
            </body>
        </html>
    );
}
