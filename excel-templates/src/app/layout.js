export default function RootLayout({ children }) {
    return (
        <html lang="en">
            <head>
                <title>Excel Filler Tool</title>
            </head>
            <body>{children}</body>
        </html>
    );
}
