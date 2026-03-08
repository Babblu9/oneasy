'use client';

import React from 'react';
import ChatPanel from '@/components/chat/ChatPanel';
import SpreadsheetPanel from '@/components/spreadsheet/SpreadsheetPanel';

export default function Home() {
    return (
        <main className="flex h-screen bg-slate-50 overflow-hidden font-sans">
            {/* Left AI Pane */}
            <section className="w-full md:w-[400px] lg:w-[450px] flex-shrink-0 z-20 shadow-2xl relative">
                <ChatPanel />
            </section>

            {/* Right Spreadsheet Pane */}
            <section className="flex-1 overflow-hidden relative z-10 hidden md:block">
                <SpreadsheetPanel />
            </section>
        </main>
    );
}
