import React, { useState, useEffect, useRef } from 'react';
import {
  Film,
  BookOpen,
  ShoppingBag,
  TrendingUp,
  Clapperboard,
  Tv,
  Coins,
  Building2,
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Minimize2,
  Download,
  RotateCcw,
  Grid,
  ExternalLink,
  Layers,
  ArrowRight,
  Calculator,
  FileSpreadsheet,
  CheckCircle2,
  Sparkles,
  LogIn,
  LogOut,
  User as UserIcon,
  Image as ImageIcon,
  Eye,
  X,
} from 'lucide-react';
import { User } from 'firebase/auth';
import { initAuth, googleSignIn, logout, getAccessToken } from './services/auth';
import { SLIDES_META, SlideMeta } from './data/slidesData';
import { GoogleSlidesModal } from './components/GoogleSlidesModal';
import { InteractiveLedgerCalculator } from './components/InteractiveLedgerCalculator';

// Generated Cinematic Visual Assets
const ASSETS = {
  hollywoodStudio: '/src/assets/images/hollywood_studio_lot_1792672084163.jpg',
  vintageLibrary: '/src/assets/images/vintage_bookstore_display_1792672084438.jpg',
  indianMultiplex: '/src/assets/images/indian_multiplex_cinema_1792672086085.jpg',
  boxOfficeWaterfall: '/src/assets/images/box_office_waterfall_1792672085369.jpg',
};

export default function App() {
  const [currentSlide, setCurrentSlide] = useState(0);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [showOverview, setShowOverview] = useState(false);
  const [showCalculator, setShowCalculator] = useState(false);
  const [showSlidesModal, setShowSlidesModal] = useState(false);
  const [isExportingPptx, setIsExportingPptx] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

  // Lightbox Modal for Fullscreen Visual Inspection
  const [activeImageZoom, setActiveImageZoom] = useState<{ src: string; caption: string; credit: string } | null>(null);

  // Authentication State
  const [user, setUser] = useState<User | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const containerRef = useRef<HTMLDivElement>(null);

  // Initialise Auth Listener
  useEffect(() => {
    const unsubscribe = initAuth(
      (currentUser, token) => {
        setUser(currentUser);
        setAccessToken(token);
      },
      () => {
        setUser(null);
        setAccessToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleGoogleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const res = await googleSignIn();
      if (res) {
        setUser(res.user);
        setAccessToken(res.accessToken);
      }
    } catch (err: any) {
      console.error('Login error:', err);
      setExportError(err.message || 'Google sign-in was interrupted or failed.');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await logout();
    setUser(null);
    setAccessToken(null);
  };

  // Keyboard Navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (showOverview || showCalculator || showSlidesModal || activeImageZoom) {
        if (e.key === 'Escape') {
          setShowOverview(false);
          setShowCalculator(false);
          setShowSlidesModal(false);
          setActiveImageZoom(null);
        }
        return;
      }
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        setCurrentSlide((prev) => Math.min(prev + 1, SLIDES_META.length - 1));
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        setCurrentSlide((prev) => Math.max(prev - 1, 0));
      } else if (e.key === 'Home') {
        e.preventDefault();
        setCurrentSlide(0);
      } else if (e.key === 'End') {
        e.preventDefault();
        setCurrentSlide(SLIDES_META.length - 1);
      } else if (e.key === 'f' || e.key === 'F') {
        toggleFullscreen();
      } else if (e.key === 'c' || e.key === 'C') {
        setShowCalculator(true);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [showOverview, showCalculator, showSlidesModal, activeImageZoom, isFullscreen]);

  // Fullscreen toggle
  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      containerRef.current?.requestFullscreen().catch((err) => {
        console.error(`Fullscreen request failed: ${err.message}`);
      });
      setIsFullscreen(true);
    } else {
      document.exitFullscreen().catch(() => {});
      setIsFullscreen(false);
    }
  };

  // PowerPoint Exporter using PptxGenJS
  const exportToPowerPoint = async () => {
    setIsExportingPptx(true);
    setExportError(null);

    try {
      if (!(window as any).PptxGenJS) {
        await new Promise<void>((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
          script.onload = () => resolve();
          script.onerror = () => reject(new Error('Unable to load PowerPoint generation library'));
          document.head.appendChild(script);
        });
      }

      const pptx = new (window as any).PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';
      pptx.author = 'Accountancy for Engineers';
      pptx.company = 'The Cinematic Ledger';
      pptx.title = 'Beyond the Box Office: Financial Analysis Case Study';

      const bgDark = { color: '101012' };
      const gold = 'C6A56B';
      const parchment = 'F0E5D0';
      const muted = 'AAA5A0';
      const burgundy = '521F2B';

      // Slide 1
      let s1 = pptx.addSlide();
      s1.background = bgDark;
      s1.addText('WARNER BROS. DISCOVERY & ENTERTAINMENT ACCOUNTING', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s1.addText('HOLLYWOOD EARNS BEYOND THE THEATRE', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      s1.addText('A film can finish its theatrical run. Its commercial life does not have to end.', { x: 0.8, y: 1.5, fontSize: 13, color: muted, italic: true });
      s1.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.1, w: 4.8, h: 2.5, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s1.addText('WARNER BROS. DISCOVERY (FY2025)\nAudited Company-Wide Figures\n\nTotal Revenue: US$ 37,296 Million\nStudios Segment: US$ 12,619 Million', { x: 1.0, y: 2.3, w: 4.4, h: 2.1, fontSize: 14, color: parchment, bold: true });
      s1.addShape(pptx.ShapeType.rect, { x: 6.0, y: 2.1, w: 6.5, h: 4.0, fill: { color: '19171B' }, line: { color: gold, width: 1 } });
      s1.addText('THE SIX CORE REVENUE STREAMS OF CINEMATIC IP', { x: 6.3, y: 2.4, fontSize: 13, color: gold, bold: true });
      const streams = [
        '1. Theatrical Box Office (Cinema Ticket Splits)',
        '2. Streaming & Subscriptions (SVOD / Max Platform)',
        '3. Content Licensing (Broadcasters, Syndication)',
        '4. Consumer Products & Brand Licensing (Merchandise)',
        '5. Video Games & Interactive (Hogwarts Legacy: $1B+ Retail)',
        '6. Home Entertainment (EST, VOD, Physical Media)'
      ];
      s1.addText(streams.join('\n\n'), { x: 6.3, y: 2.8, w: 6.0, h: 3.1, fontSize: 11, color: parchment });
      s1.addText('Source: SEC EDGAR WBD FY2025 Form 10-K | Note: Segment numbers represent multi-franchise totals, not HP-only.', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 2
      let s2 = pptx.addSlide();
      s2.background = bgDark;
      s2.addText('PUBLISHING REVENUE RECOGNITION', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s2.addText('HARRY POTTER: THE POWER OF BOOK SALES', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      const bookTable = [
        [{ text: 'Financial Year', options: { bold: true, color: gold } }, { text: 'Scholastic HP Trade Revenue', options: { bold: true, color: gold } }, { text: 'Key Operational Catalyst', options: { bold: true, color: gold } }],
        ['FY2006 (Ended May 31)', 'US$ 195.0 Million', 'Post Half-Blood Prince momentum'],
        ['FY2007 (Ended May 31)', 'US$ 20.0 Million', 'Interim trough prior to series finale'],
        ['FY2008 (Ended May 31)', 'US$ 270.0 Million', 'Deathly Hallows release (July 2007)'],
      ];
      s2.addTable(bookTable, { x: 0.8, y: 2.0, w: 7.0, fill: { color: '19171B' }, color: parchment, fontSize: 11, border: { pt: 1, color: gold } });
      s2.addShape(pptx.ShapeType.rect, { x: 8.2, y: 2.0, w: 4.3, h: 4.2, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s2.addText('ILLUSTRATIVE JOURNAL ENTRY\nBook Sales under ASC 606\n\nDr. Cash / Accounts Receivable  $270M\n    Cr. Sales Revenue (Book Trade)  $270M\n\n[Recording sales upon transfer of control, net of estimated return provisions.]', { x: 8.4, y: 2.2, w: 3.9, h: 3.8, fontSize: 11, color: parchment });
      s2.addText('Source: SEC EDGAR Scholastic Corp FY2008 10-K (0000930413-08-004578).', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 3
      let s3 = pptx.addSlide();
      s3.background = bgDark;
      s3.addText('RETAIL INVENTORY & LICENSING DUALITY', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s3.addText('MERCHANDISE: MINISO × HARRY POTTER × KINDER JOY', { x: 0.8, y: 0.95, fontSize: 24, color: parchment, bold: true });
      s3.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.8, w: 3.6, h: 4.5, fill: { color: '19171B' }, line: { color: gold, width: 1 } });
      s3.addText('PANEL A: MINISO RETAIL\n\n• 800+ SKU Global Collection\n• Pop-up generated >HK$5M in 7 days\n• MINISO recognizes Gross Retail Sales\n• Pays royalty percentage to WBD\n• Demonstrates lifestyle retail conversion', { x: 0.95, y: 1.95, w: 3.3, h: 4.2, fontSize: 10.5, color: parchment });
      s3.addShape(pptx.ShapeType.rect, { x: 4.8, y: 1.8, w: 3.6, h: 4.5, fill: { color: '19171B' }, line: { color: gold, width: 1 } });
      s3.addText('PANEL B: FERRERO KINDER JOY\n\n• High-velocity FMCG tie-in\n• Collectible toy figurines in every egg\n• Indian retail frenzy & secondary trade\n• Ferrero books product sales\n• WBD records licensing royalty income\n• Note: Secondary estimates are non-audited', { x: 4.95, y: 1.95, w: 3.3, h: 4.2, fontSize: 10.5, color: parchment });
      s3.addShape(pptx.ShapeType.rect, { x: 8.8, y: 1.8, w: 3.8, h: 4.5, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s3.addText('PANEL C: ACCOUNTING FLOW\n\n1. On Inventory Purchase:\nDr. Inventory | Cr. Trade Payables\n\n2. On Product Sale:\nDr. Cash / Receivables | Cr. Sales Revenue\n\n3. On COGS Recognition:\nDr. Cost of Goods Sold | Cr. Inventory\n\nFormula: Gross Profit = Sales - COGS', { x: 8.95, y: 1.95, w: 3.5, h: 4.2, fontSize: 10.5, color: parchment });
      s3.addText('Sources: MINISO Brand Announcement (2024); IIT Madras Article 19 (Feb 2025).', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 4
      let s4 = pptx.addSlide();
      s4.background = bgDark;
      s4.addText('EXPENSE RECOGNITION & CONTRACTUAL ROYALTIES', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s4.addText('ROYALTIES: THE COST OF BORROWING A STORY', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      const royaltyTable = [
        [{ text: 'Scholastic Metric', options: { bold: true, color: gold } }, { text: 'FY2016', options: { bold: true, color: gold } }, { text: 'FY2017', options: { bold: true, color: gold } }, { text: 'Variance', options: { bold: true, color: gold } }],
        ['Royalty Expense', 'US$ 92.0 M', 'US$ 146.4 M', '+US$ 54.4 M (+59.1%)'],
        ['Total Cost of Goods Sold', 'US$ 762.3 M', 'US$ 814.5 M', '+US$ 52.2 M (+6.8%)'],
        ["Children's Book Publishing Op. Inc.", 'US$ 120.6 M', 'US$ 143.1 M', '+US$ 22.5 M (+18.7%)'],
      ];
      s4.addTable(royaltyTable, { x: 0.8, y: 2.0, w: 7.2, fill: { color: '19171B' }, color: parchment, fontSize: 10.5, border: { pt: 1, color: gold } });
      s4.addShape(pptx.ShapeType.rect, { x: 8.4, y: 2.0, w: 4.2, h: 4.2, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s4.addText('PUBLISHER JOURNAL ENTRY\n\nDr. Royalty Expense  $54.4M\n    Cr. Royalty Payable  $54.4M\n\nPublisher records royalty as cost of sales;\nRights holder (Author/Studio) records Royalty Revenue upon satisfaction of contract performance obligations.', { x: 8.6, y: 2.2, w: 3.8, h: 3.8, fontSize: 11, color: parchment });
      s4.addText('Source: Scholastic Form 10-K for period ended May 31, 2017 (Commission File 000-0866729).', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 5
      let s5 = pptx.addSlide();
      s5.background = bgDark;
      s5.addText('FILM ASSET CAPITALISATION (US GAAP ASC 926)', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s5.addText('FILM PRODUCTION COSTS & AMORTISATION', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      const wbdCostTable = [
        [{ text: 'Content Classification Category', options: { bold: true, color: gold } }, { text: 'Carrying Value (US$ Millions)', options: { bold: true, color: gold } }],
        ['Released production costs, less amortisation', '$ 8,692 M'],
        ['Completed but not released', '$ 1,630 M'],
        ['In production and other development', '$ 4,326 M'],
        ['Total Warner Bros. Discovery Production Costs', '$ 14,648 M'],
      ];
      s5.addTable(wbdCostTable, { x: 0.8, y: 2.0, w: 7.2, fill: { color: '19171B' }, color: parchment, fontSize: 11, border: { pt: 1, color: gold } });
      s5.addShape(pptx.ShapeType.rect, { x: 8.4, y: 2.0, w: 4.2, h: 4.2, fill: { color: '19171B' }, line: { color: gold, width: 1 } });
      s5.addText('ACCOUNTING LIFECYCLE\n\n1. Production Spending:\nCapitalised as an Asset (not immediate expense).\n\n2. Theatrical / Digital Release:\nAsset amortised matching revenue recognition.\n\n3. Illustrative Entry:\nDr. Content Amortisation Expense\n    Cr. Film / Content Asset', { x: 8.6, y: 2.2, w: 3.8, h: 3.8, fontSize: 11, color: parchment });
      s5.addText('Source: Warner Bros. Discovery Form 10-K (FY2025). Company-wide figures, not limited to one franchise.', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 6
      let s6 = pptx.addSlide();
      s6.background = bgDark;
      s6.addText('INDIAN THEATRICAL EXHIBITION ECOSYSTEM', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s6.addText('INDIAN CINEMA: BEYOND THE TICKET COUNTER', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      s6.addText('VALUE CHAIN: FILM PRODUCER → DISTRIBUTOR → CINEMA EXHIBITOR (PVR INOX) → PATRON', { x: 0.8, y: 1.8, fontSize: 13, color: gold, bold: true });
      const pvrStreams = [
        '1. Box Office Net Revenue: Tickets net of municipal entertainment tax & GST.',
        '2. Food & Beverage (F&B): High gross-margin engine (often 70-75% gross margin).',
        '3. On-Screen & In-Cinema Advertising: High EBITDA margin marketing channel.',
        '4. Convenience Fees & Concessions: Digital ticketing partnership commissions.'
      ];
      s6.addText(pvrStreams.join('\n\n'), { x: 0.8, y: 2.5, w: 6.5, h: 3.5, fontSize: 12, color: parchment });
      s6.addShape(pptx.ShapeType.rect, { x: 7.8, y: 2.4, w: 4.7, h: 3.8, fill: { color: '19171B' }, line: { color: gold, width: 1 } });
      s6.addText('EXHIBITOR REVENUE DYNAMICS\n\nGross Box Office is NOT Exhibitor Revenue.\nGross collections include GST (12% or 18%).\nNet Box Office is split between Distributor and Exhibitor (typically 50:50 Week 1 sliding to 60:40).\n\nF&B and Advertising provide operating resilience when theatrical box office fluctuates.', { x: 8.0, y: 2.6, w: 4.3, h: 3.4, fontSize: 11, color: parchment });
      s6.addText('Source: PVR INOX Ltd Annual Report FY2024-25 (NSE Filing PVR_14082025115149).', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 7
      let s7 = pptx.addSlide();
      s7.background = bgDark;
      s7.addText('INDIAN ACCOUNTING STANDARDS (IND AS) ANALYSIS', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s7.addText('PVR INOX: INSIDE THE BALANCE SHEET', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      const pvrBalance = [
        [{ text: 'Balance Sheet Item (Consolidated)', options: { bold: true, color: gold } }, { text: 'Amount (₹ Millions)', options: { bold: true, color: gold } }],
        ['Equity Share Capital', '₹ 982 M'],
        ['Other Equity (Reserves & Surplus)', '₹ 69,265 M'],
        ["Total Equity (Shareholders' Funds)", '₹ 70,247 M'],
        ['Long-Term Borrowings', '₹ 9,198 M'],
        ['Current Liabilities', '₹ 24,977 M'],
        ['Non-Current Assets (inc. Right-of-Use Leases)', '₹ 145,920 M'],
        ['Current Assets', '₹ 10,810 M'],
        ['Total Assets = Total Equity & Liabilities', '₹ 156,730 M'],
      ];
      s7.addTable(pvrBalance, { x: 0.8, y: 1.8, w: 7.2, fill: { color: '19171B' }, color: parchment, fontSize: 10, border: { pt: 1, color: gold } });
      s7.addShape(pptx.ShapeType.rect, { x: 8.4, y: 1.8, w: 4.2, h: 4.5, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s7.addText('KEY BALANCE SHEET LESSONS\n\n1. Ind AS 116 Lease Capitalisation:\nCinema exhibitors hold huge Right-of-Use assets and corresponding lease liabilities.\n\n2. Capital Intensive Model:\nTotal Assets of ₹156,730M reflects multiplex buildouts, projectors, screens, and long-term leaseholds.\n\n3. Balance Sheet ≠ Profit:\nSolvency does not guarantee annual operating net profitability.', { x: 8.6, y: 2.0, w: 3.8, h: 4.1, fontSize: 10.5, color: parchment });
      s7.addText('Source: PVR INOX Audited Consolidated Financial Statements as of 31 March 2025.', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 8
      let s8 = pptx.addSlide();
      s8.background = bgDark;
      s8.addText('COMMERCIAL REALITY OF BOX OFFICE REPORTING', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s8.addText('DHURANDHAR: BOX OFFICE IS NOT PRODUCER PROFIT', { x: 0.8, y: 0.95, fontSize: 24, color: parchment, bold: true });
      const dhurandharTable = [
        [{ text: 'Waterfall Dimension', options: { bold: true, color: gold } }, { text: 'Estimated Value', options: { bold: true, color: gold } }, { text: 'Accounting Nature', options: { bold: true, color: gold } }],
        ['India Gross Collections', '₹ 950 Crore', 'Gross box office ticket receipts'],
        ['Estimated India Theatre Share', '₹ 300 Crore', 'Retained by multiplex & single screen owners'],
        ['Worldwide Gross Box Office', '₹ 1,250 Crore', 'Global theatrical gross (pre-deductions)'],
      ];
      s8.addTable(dhurandharTable, { x: 0.8, y: 1.8, w: 7.4, fill: { color: '19171B' }, color: parchment, fontSize: 10.5, border: { pt: 1, color: gold } });
      s8.addShape(pptx.ShapeType.rect, { x: 8.6, y: 1.8, w: 4.0, h: 4.4, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s8.addText('THE PRODUCER PROFIT FORMULA\n\nPRODUCER PROFIT =\nRecognised Producer Share\n- Production Costs\n- P&A (Prints & Advertising)\n- Financing & Distribution Margins\n\nAudited producer accounts are not public. Exact producer profit cannot be reliably calculated from gross box-office collections alone.', { x: 8.8, y: 2.0, w: 3.6, h: 4.0, fontSize: 10.5, color: parchment });
      s8.addText('Source: The Week Box-Office Financial Analysis (Feb 3, 2026). Figures are industry estimates.', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 9
      let s9 = pptx.addSlide();
      s9.background = bgDark;
      s9.addText('COMPARATIVE FINANCIAL FRAMEWORK', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s9.addText('HOLLYWOOD VS INDIAN CINEMA: AN ACCOUNTING MATRIX', { x: 0.8, y: 0.95, fontSize: 24, color: parchment, bold: true });
      const matrixTable = [
        [{ text: 'Dimension', options: { bold: true, color: gold } }, { text: 'Hollywood Studio (e.g. WBD)', options: { bold: true, color: gold } }, { text: 'Indian Cinema Exhibitor (e.g. PVR INOX)', options: { bold: true, color: gold } }],
        ['Core Revenue Streams', 'Content distribution, SVOD, IP licensing, games', 'Box office ticket admissions, F&B, advertising'],
        ['Licensing & Royalties', 'Primary driver of perpetual high-margin profit', 'Incidental / minor operational element'],
        ['Production Costs', 'Capitalised on Balance Sheet; amortised over time', 'Does not produce films; incurs film hire share'],
        ['Merchandise Inventory', 'Licensed to third parties (MINISO, Kinder Joy)', 'Carries F&B inventory (popcorn, beverages)'],
        ['Primary Asset Base', 'Intangible IP rights, film library ($14.6B+ assets)', 'Tangible screens, cinema leases, audio/projectors'],
        ['Financial Goal', 'Maximise franchise lifetime cash flows globally', 'Maximise spend per head (SPH) & seat occupancy'],
      ];
      s9.addTable(matrixTable, { x: 0.8, y: 1.8, w: 11.6, fill: { color: '19171B' }, color: parchment, fontSize: 10, border: { pt: 1, color: gold } });
      s9.addText('Note: Studio and Exhibitor models occupy distinct positions on the media value chain and are not directly comparable without adjusting for scope.', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      // Slide 10
      let s10 = pptx.addSlide();
      s10.background = bgDark;
      s10.addText('EXECUTIVE SYNTHESIS & STATUTORY SOURCES', { x: 0.8, y: 0.6, fontSize: 11, color: gold, bold: true });
      s10.addText('THE FINANCIAL TAKEAWAY & AUDIT LEDGER', { x: 0.8, y: 0.95, fontSize: 26, color: parchment, bold: true });
      s10.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.8, w: 5.6, h: 4.5, fill: { color: burgundy }, line: { color: gold, width: 1 } });
      s10.addText('THREE GOLDEN RULES OF ENTERTAINMENT ACCOUNTING\n\n1. Analyse Revenue by Stream\nNever confuse box office collections with studio revenue or exhibitor income.\n\n2. Track Capitalisation & Costs\nDistinguish amortisation of content assets from immediate operational expenditure.\n\n3. Look Beyond Gross Metrics\nTrue equity value is proven through the Balance Sheet and Cash Flow Statement, not viral headlines.', { x: 1.0, y: 2.0, w: 5.2, h: 4.1, fontSize: 11, color: parchment });
      s10.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.8, w: 5.6, h: 4.5, fill: { color: '19171B' }, line: { color: gold, width: 1 } });
      s10.addText('PRIMARY REGULATORY & STATUTORY ARCHIVE\n\n1. SEC EDGAR: Scholastic FY2008 & FY2017 Form 10-K\n2. SEC EDGAR: Warner Bros. Discovery FY2025 Form 10-K\n3. NSE India: PVR INOX FY2024-25 Consolidated Report\n4. WIPO: Harry Potter and the IP Bonanza\n5. MCA India: Ind AS 115 Revenue from Contracts\n6. The Week: Dhurandhar Theatrical Box-Office Audit', { x: 7.0, y: 2.0, w: 5.2, h: 4.1, fontSize: 10.5, color: parchment });
      s10.addText('All figures cited are drawn from audited regulatory filings or formally designated industry estimates.', { x: 0.8, y: 6.7, fontSize: 9, color: muted });

      await pptx.writeFile({ fileName: 'Beyond_the_Box_Office_Accounting_Case_Study.pptx' });
      setIsExportingPptx(false);
    } catch (err: any) {
      console.error(err);
      setExportError(err.message || 'Failed to generate PowerPoint file.');
      setIsExportingPptx(false);
    }
  };

  const activeMeta = SLIDES_META[currentSlide];

  return (
    <div
      ref={containerRef}
      className="relative w-full h-screen bg-[#101012] text-[#F0E5D0] font-sans flex flex-col justify-between overflow-hidden select-none"
      style={{
        backgroundImage: 'radial-gradient(ellipse at 50% 20%, #1c181f 0%, #101012 80%)',
      }}
    >
      {/* Top Header Bar */}
      <header className="flex items-center justify-between px-6 md:px-8 py-3.5 border-b border-[#C6A56B]/20 bg-[#101012]/80 backdrop-blur-md z-30">
        <div className="flex items-center space-x-3">
          <div className="flex items-center space-x-2 text-[#C6A56B]">
            <Clapperboard className="w-5 h-5 text-[#C6A56B]" />
            <span className="font-serif-heading tracking-widest text-xs uppercase font-bold text-[#C6A56B]">
              The Cinematic Ledger
            </span>
          </div>
          <span className="text-[#AAA5A0]/40">|</span>
          <span className="text-xs tracking-wider text-[#AAA5A0] hidden lg:inline font-serif-body">
            Accountancy for Engineers · Hollywood &amp; Indian Cinema Case Study
          </span>
        </div>

        {/* Action Controls & Google Account */}
        <div className="flex items-center space-x-2.5">
          {/* Simulator Button */}
          <button
            onClick={() => setShowCalculator(true)}
            className="flex items-center space-x-1.5 px-3 py-1.5 rounded text-xs text-[#F0E5D0] bg-[#19171B] border border-[#C6A56B]/30 hover:border-[#C6A56B] transition shadow-sm"
            title="Open Interactive Financial Simulator (Press C)"
          >
            <Calculator className="w-3.5 h-3.5 text-[#C6A56B]" />
            <span className="hidden sm:inline">Ledger Simulator</span>
          </button>

          {/* Slide Overview Grid Button */}
          <button
            onClick={() => setShowOverview(!showOverview)}
            className="flex items-center space-x-1.5 px-3 py-1.5 rounded text-xs text-[#F0E5D0] bg-[#19171B] border border-[#C6A56B]/30 hover:border-[#C6A56B] transition"
            title="Slide Overview (Grid)"
          >
            <Grid className="w-3.5 h-3.5 text-[#C6A56B]" />
            <span className="hidden md:inline">Overview</span>
          </button>

          {/* Export to Google Slides Button */}
          <button
            onClick={() => setShowSlidesModal(true)}
            className="flex items-center space-x-1.5 px-3.5 py-1.5 rounded text-xs font-semibold text-[#101012] bg-[#F4B400] hover:bg-[#ffc933] transition shadow-md"
            title="Export 10 Slides directly to your Google Slides account"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none">
              <path
                d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10H7v-2h10v2zm0-4H7V7h10v2z"
                fill="#101012"
              />
            </svg>
            <span>Google Slides</span>
          </button>

          {/* Download PowerPoint (.pptx) Button */}
          <button
            onClick={exportToPowerPoint}
            disabled={isExportingPptx}
            className="flex items-center space-x-1.5 px-3 py-1.5 rounded text-xs font-semibold text-[#101012] bg-[#C6A56B] hover:bg-[#e0c38c] transition shadow-md disabled:opacity-50"
            title="Download Editable 10-Slide .pptx"
          >
            <Download className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">{isExportingPptx ? 'Exporting...' : 'Export PPTX'}</span>
          </button>

          {/* Fullscreen Toggle */}
          <button
            onClick={toggleFullscreen}
            className="p-1.5 rounded text-[#AAA5A0] hover:text-[#C6A56B] bg-[#19171B] border border-[#C6A56B]/20 transition"
            title="Toggle Fullscreen (F)"
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>

          {/* Google Account Profile / Sign-In Button */}
          <div className="pl-2 border-l border-[#C6A56B]/20">
            {user ? (
              <div className="flex items-center space-x-2 bg-[#19171B] border border-[#C6A56B]/30 rounded-full py-0.5 px-2.5">
                {user.photoURL ? (
                  <img
                    src={user.photoURL}
                    alt={user.displayName || 'Google User'}
                    referrerPolicy="no-referrer"
                    className="w-5 h-5 rounded-full border border-[#C6A56B]"
                  />
                ) : (
                  <UserIcon className="w-4 h-4 text-[#C6A56B]" />
                )}
                <span className="text-[11px] font-sans text-[#F0E5D0] max-w-[90px] truncate hidden sm:inline">
                  {user.displayName?.split(' ')[0] || user.email?.split('@')[0]}
                </span>
                <button
                  onClick={handleLogout}
                  className="p-1 text-[#AAA5A0] hover:text-red-400 transition"
                  title="Sign out of Google"
                >
                  <LogOut className="w-3 h-3" />
                </button>
              </div>
            ) : (
              <button
                onClick={handleGoogleLogin}
                disabled={isLoggingIn}
                className="flex items-center space-x-1.5 px-2.5 py-1.5 rounded text-[11px] text-[#F0E5D0] bg-[#19171B] border border-[#C6A56B]/30 hover:border-[#C6A56B] transition"
                title="Sign in with Google to enable Google Slides"
              >
                <svg className="w-3.5 h-3.5" viewBox="0 0 48 48">
                  <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z" />
                  <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z" />
                  <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z" />
                  <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z" />
                </svg>
                <span className="hidden sm:inline">Sign in</span>
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Main Slide Stage (16:9 Aspect Presentation Frame) */}
      <main className="relative flex-1 flex items-center justify-center p-3 sm:p-5 lg:p-7 overflow-hidden z-10">
        <div className="w-full max-w-7xl aspect-[16/9] max-h-full bg-[#19171B]/95 border border-[#C6A56B]/30 rounded-xl shadow-2xl flex flex-col justify-between p-4 sm:p-6 md:p-8 relative backdrop-blur-xl overflow-hidden transition-all duration-300">
          {/* Subtle Ledger Grid Background Accent */}
          <div
            className="absolute inset-0 pointer-events-none opacity-5"
            style={{
              backgroundImage:
                'linear-gradient(#C6A56B 1px, transparent 1px), linear-gradient(90deg, #C6A56B 1px, transparent 1px)',
              backgroundSize: '32px 32px',
            }}
          />

          {/* Slide Heading Area */}
          <div className="relative z-10">
            <div className="flex items-center justify-between mb-1.5">
              <span className="text-[11px] font-semibold tracking-widest text-[#C6A56B] uppercase flex items-center gap-2">
                <span className="inline-block w-2 h-2 rounded-full bg-[#C6A56B]" />
                {activeMeta.category}
              </span>
              <span className="font-ledger text-sm text-[#C6A56B]/70 tracking-wider">
                {activeMeta.slideNumber} / 10
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl md:text-3xl font-serif-heading text-[#F0E5D0] tracking-tight font-bold">
              {activeMeta.title}
            </h1>
            <p className="text-xs sm:text-sm text-[#AAA5A0] italic mt-0.5 font-serif-body">
              "{activeMeta.headline}"
            </p>
          </div>

          {/* Slide Interactive Content Canvas with Images */}
          <div className="relative z-10 flex-1 my-2.5 flex items-center justify-center overflow-hidden">
            {renderSlideContent(currentSlide, (src, caption, credit) =>
              setActiveImageZoom({ src, caption, credit })
            )}
          </div>

          {/* Slide Footer / Source Verification Panel */}
          <div className="relative z-10 pt-2.5 border-t border-[#C6A56B]/20 flex flex-col sm:flex-row items-start sm:items-center justify-between text-[11px] text-[#AAA5A0]">
            <div className="flex items-center gap-2">
              <span className="font-bold text-[#C6A56B] uppercase tracking-wider text-[10px]">
                Statutory Source:
              </span>
              <span className="truncate max-w-md">{activeMeta.sourceNote}</span>
            </div>
            <div className="flex items-center gap-3 mt-1 sm:mt-0">
              <span className="text-[10px] text-[#829C80] font-ledger">Verified Ind AS / ASC 606</span>
              <a
                href={activeMeta.sourceUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center gap-1 text-[#C6A56B] hover:underline"
              >
                Inspect Official Filing <ExternalLink className="w-3 h-3" />
              </a>
            </div>
          </div>
        </div>
      </main>

      {/* Slide Navigation & Progress Bar */}
      <footer className="px-6 md:px-8 py-3 bg-[#101012] border-t border-[#C6A56B]/20 flex items-center justify-between z-30">
        <div className="flex items-center space-x-2">
          <button
            onClick={() => setCurrentSlide((prev) => Math.max(prev - 1, 0))}
            disabled={currentSlide === 0}
            className="p-2 rounded border border-[#C6A56B]/30 text-[#F0E5D0] hover:bg-[#C6A56B]/10 disabled:opacity-30 disabled:hover:bg-transparent transition"
            title="Previous Slide (←)"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>
          <span className="text-xs text-[#AAA5A0] font-ledger px-2">
            SLIDE {currentSlide + 1} OF 10
          </span>
          <button
            onClick={() => setCurrentSlide((prev) => Math.min(prev + 1, SLIDES_META.length - 1))}
            disabled={currentSlide === SLIDES_META.length - 1}
            className="p-2 rounded border border-[#C6A56B]/30 text-[#F0E5D0] hover:bg-[#C6A56B]/10 disabled:opacity-30 disabled:hover:bg-transparent transition"
            title="Next Slide (→)"
          >
            <ChevronRight className="w-4 h-4" />
          </button>
        </div>

        {/* Dynamic Progress Indicator */}
        <div className="flex-1 max-w-md mx-6 hidden sm:block">
          <div className="w-full bg-[#19171B] h-1.5 rounded-full overflow-hidden border border-[#C6A56B]/20">
            <div
              className="bg-[#C6A56B] h-full transition-all duration-300"
              style={{ width: `${((currentSlide + 1) / SLIDES_META.length) * 100}%` }}
            />
          </div>
        </div>

        <div className="flex items-center space-x-4 text-xs text-[#AAA5A0]">
          <span className="hidden lg:inline text-[11px]">Use [←] / [→] or Space to navigate</span>
          <button
            onClick={() => setCurrentSlide(0)}
            className="flex items-center gap-1 hover:text-[#C6A56B] transition"
            title="Restart Presentation (Home)"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Restart</span>
          </button>
        </div>
      </footer>

      {/* Lightbox Image Zoom Modal */}
      {activeImageZoom && (
        <div
          className="fixed inset-0 z-50 bg-[#101012]/95 backdrop-blur-md flex items-center justify-center p-4 md:p-8"
          onClick={() => setActiveImageZoom(null)}
        >
          <div
            className="relative max-w-5xl w-full bg-[#19171B] border border-[#C6A56B]/50 rounded-xl overflow-hidden shadow-2xl flex flex-col"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between p-3.5 border-b border-[#C6A56B]/20 bg-[#101012]">
              <span className="font-serif-heading text-xs text-[#C6A56B] font-bold uppercase tracking-wider">
                {activeImageZoom.credit}
              </span>
              <button
                onClick={() => setActiveImageZoom(null)}
                className="p-1 rounded text-[#AAA5A0] hover:text-[#F0E5D0] hover:bg-[#521F2B]/50 transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="relative aspect-[16/9] w-full bg-black overflow-hidden">
              <img
                src={activeImageZoom.src}
                alt={activeImageZoom.caption}
                referrerPolicy="no-referrer"
                className="w-full h-full object-cover"
              />
            </div>
            <div className="p-4 bg-[#101012] border-t border-[#C6A56B]/20 flex items-center justify-between text-xs text-[#F0E5D0]">
              <span className="font-serif-body italic">{activeImageZoom.caption}</span>
              <span className="text-[10px] text-[#C6A56B] font-ledger">Press ESC to close</span>
            </div>
          </div>
        </div>
      )}

      {/* Slide Overview Grid Modal with Thumbnail Previews */}
      {showOverview && (
        <div className="fixed inset-0 z-50 bg-[#101012]/95 backdrop-blur-md flex flex-col p-6 md:p-8 overflow-y-auto">
          <div className="flex justify-between items-center mb-6 max-w-6xl mx-auto w-full">
            <div className="flex items-center gap-2">
              <Grid className="w-5 h-5 text-[#C6A56B]" />
              <h2 className="text-xl font-serif-heading text-[#C6A56B]">
                Presentation Index · 10 Widescreen Slides with Visual Ledger
              </h2>
            </div>
            <button
              onClick={() => setShowOverview(false)}
              className="px-4 py-1.5 text-xs bg-[#521F2B] text-white rounded hover:bg-[#6b2938] transition font-semibold"
            >
              Close
            </button>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4 max-w-6xl mx-auto w-full">
            {SLIDES_META.map((s, idx) => {
              // Map representative image for preview
              let thumbImage = ASSETS.hollywoodStudio;
              if (idx === 1 || idx === 3) thumbImage = ASSETS.vintageLibrary;
              if (idx === 5 || idx === 6) thumbImage = ASSETS.indianMultiplex;
              if (idx === 7 || idx === 9) thumbImage = ASSETS.boxOfficeWaterfall;

              return (
                <div
                  key={s.id}
                  onClick={() => {
                    setCurrentSlide(idx);
                    setShowOverview(false);
                  }}
                  className={`p-3 rounded-lg cursor-pointer border transition flex flex-col justify-between h-44 group ${
                    currentSlide === idx
                      ? 'border-[#C6A56B] bg-[#521F2B]/40 shadow-xl ring-1 ring-[#C6A56B]'
                      : 'border-[#C6A56B]/20 bg-[#19171B] hover:border-[#C6A56B]/60'
                  }`}
                >
                  <div className="space-y-2">
                    <div className="flex justify-between items-center">
                      <span className="text-[10px] text-[#C6A56B] font-ledger font-bold">
                        SLIDE {s.slideNumber}
                      </span>
                      <ImageIcon className="w-3 h-3 text-[#AAA5A0]/60 group-hover:text-[#C6A56B]" />
                    </div>
                    {/* Thumbnail banner */}
                    <div className="w-full h-14 rounded overflow-hidden relative border border-[#C6A56B]/20">
                      <img
                        src={thumbImage}
                        alt={s.title}
                        referrerPolicy="no-referrer"
                        className="w-full h-full object-cover group-hover:scale-105 transition duration-300 brightness-90"
                      />
                    </div>
                    <p className="text-xs font-serif-heading font-semibold text-[#F0E5D0] line-clamp-1">
                      {s.title}
                    </p>
                  </div>
                  <span className="text-[9px] text-[#AAA5A0] line-clamp-1 italic font-serif-body">
                    {s.category}
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Google Slides Export & Creation Modal */}
      <GoogleSlidesModal
        isOpen={showSlidesModal}
        onClose={() => setShowSlidesModal(false)}
        accessToken={accessToken}
        onRequireAuth={handleGoogleLogin}
      />

      {/* Interactive Financial Simulator Drawer / Modal */}
      <InteractiveLedgerCalculator
        isOpen={showCalculator}
        onClose={() => setShowCalculator(false)}
      />

      {/* Export Error Alert */}
      {exportError && (
        <div className="fixed bottom-16 right-8 bg-[#521F2B] border border-red-500 text-white text-xs px-4 py-3 rounded-lg shadow-xl z-50 flex items-center gap-3">
          <span>{exportError}</span>
          <button
            onClick={() => setExportError(null)}
            className="underline font-bold text-[#F0E5D0]"
          >
            Dismiss
          </button>
        </div>
      )}
    </div>
  );
}

// ==========================================
// INDIVIDUAL SLIDE CONTENT RENDERERS
// ==========================================

function renderSlideContent(
  slideIndex: number,
  onZoom: (src: string, caption: string, credit: string) => void
) {
  switch (slideIndex) {
    case 0:
      return <Slide1Content onZoom={onZoom} />;
    case 1:
      return <Slide2Content onZoom={onZoom} />;
    case 2:
      return <Slide3Content onZoom={onZoom} />;
    case 3:
      return <Slide4Content onZoom={onZoom} />;
    case 4:
      return <Slide5Content onZoom={onZoom} />;
    case 5:
      return <Slide6Content onZoom={onZoom} />;
    case 6:
      return <Slide7Content onZoom={onZoom} />;
    case 7:
      return <Slide8Content onZoom={onZoom} />;
    case 8:
      return <Slide9Content onZoom={onZoom} />;
    case 9:
      return <Slide10Content onZoom={onZoom} />;
    default:
      return null;
  }
}

// SLIDE 1: HOLLYWOOD REVENUE STREAMS & WBD OVERVIEW (WITH STUDIO IMAGE)
function Slide1Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  const streams = [
    { title: 'Theatrical Box Office', desc: 'First window theatrical exhibition ticket revenues' },
    { title: 'Streaming & SVOD', desc: 'Digital platform fees (e.g. Max global subscriptions)' },
    { title: 'Content Licensing', desc: 'Syndication rights to TV broadcasters & streamers' },
    { title: 'Merchandise & IP', desc: 'Consumer products, apparel, toys, and licensing' },
    { title: 'Video Games', desc: 'Interactive entertainment (e.g. Hogwarts Legacy: $1B+)' },
    { title: 'Home Entertainment', desc: 'Electronic sell-through (EST), VOD & physical sales' },
  ];

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-center">
      {/* Left Data Highlight Panel with Hollywood Soundstage Thumbnail */}
      <div className="lg:col-span-5 bg-[#521F2B]/40 border border-[#C6A56B]/30 rounded-lg p-4 flex flex-col justify-between h-full">
        <div>
          <div className="flex items-center justify-between mb-1">
            <span className="text-[10px] font-ledger uppercase text-[#C6A56B] tracking-wider block">
              Audited Corporate Disclosures · FY2025
            </span>
            <button
              onClick={() =>
                onZoom(
                  ASSETS.hollywoodStudio,
                  'Historic Hollywood soundstage and movie backlot exhibition — Warner Bros. Discovery studio lot.',
                  'STUDIO PRODUCTION INFRASTRUCTURE'
                )
              }
              className="flex items-center gap-1 text-[10px] text-[#C6A56B] hover:underline"
              title="Inspect high-res studio soundstage"
            >
              <Eye className="w-3 h-3" />
              <span>Inspect Lot</span>
            </button>
          </div>

          <h3 className="text-lg font-serif-heading text-[#F0E5D0] font-bold">Warner Bros. Discovery</h3>

          {/* Integrated Image Frame */}
          <div
            onClick={() =>
              onZoom(
                ASSETS.hollywoodStudio,
                'Hollywood Soundstage 12 & Projection Beam: The physical infrastructure where cinematic IP is born.',
                'WARNER BROS. PRODUCTION ASSETS'
              )
            }
            className="my-2.5 h-24 w-full rounded-md overflow-hidden relative border border-[#C6A56B]/30 cursor-pointer group shadow-md"
          >
            <img
              src={ASSETS.hollywoodStudio}
              alt="Hollywood Movie Soundstage Backlot"
              referrerPolicy="no-referrer"
              className="w-full h-full object-cover group-hover:scale-105 transition duration-500 brightness-95"
            />
            <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-transparent to-transparent flex items-end p-2">
              <span className="text-[10px] text-[#F0E5D0] font-serif-body italic flex items-center gap-1">
                <ImageIcon className="w-3 h-3 text-[#C6A56B]" /> Soundstage 12 · Warner Bros. Production Lot
              </span>
            </div>
          </div>

          <div className="space-y-2 mt-2">
            <div className="bg-[#101012]/85 border-l-4 border-[#C6A56B] p-2.5 rounded">
              <span className="text-[9px] text-[#AAA5A0] uppercase font-ledger">Total Company Revenue (FY2025)</span>
              <p className="text-xl font-ledger text-[#F0E5D0] font-bold">US$ 37,296 Million</p>
              <span className="text-[9px] text-[#829C80]">Consolidated Warner Bros. Discovery total</span>
            </div>

            <div className="bg-[#101012]/85 border-l-4 border-[#829C80] p-2.5 rounded">
              <span className="text-[9px] text-[#AAA5A0] uppercase font-ledger">Studios Segment Revenue (FY2025)</span>
              <p className="text-xl font-ledger text-[#F0E5D0] font-bold">US$ 12,619 Million</p>
              <span className="text-[9px] text-[#AAA5A0]">Film, TV production &amp; content licensing</span>
            </div>
          </div>
        </div>

        <div className="mt-2 pt-2 border-t border-[#C6A56B]/20 text-[9px] text-[#AAA5A0]">
          *Note: These are WBD company-wide and segment-level figures, not Harry Potter-specific totals.
        </div>
      </div>

      {/* Right Stream Radial / Grid Panel */}
      <div className="lg:col-span-7 flex flex-col justify-between h-full space-y-2.5">
        <div className="text-xs font-serif-heading uppercase tracking-wider text-[#C6A56B] font-bold">
          The 6 Revenue Pathways of Modern Cinematic Intellectual Property
        </div>
        <div className="grid grid-cols-2 gap-2.5">
          {streams.map((s, idx) => (
            <div
              key={idx}
              className="bg-[#101012]/75 border border-[#C6A56B]/20 rounded-md p-2.5 hover:border-[#C6A56B]/60 transition"
            >
              <div className="flex items-center gap-2 mb-1">
                <span className="w-5 h-5 rounded-full bg-[#521F2B] text-[#C6A56B] text-[10px] flex items-center justify-center font-bold font-ledger">
                  0{idx + 1}
                </span>
                <span className="text-xs font-semibold text-[#F0E5D0]">{s.title}</span>
              </div>
              <p className="text-[10px] text-[#AAA5A0] leading-tight font-serif-body">{s.desc}</p>
            </div>
          ))}
        </div>

        {/* Flow Strip */}
        <div className="p-2 bg-[#101012]/90 border border-[#C6A56B]/30 rounded text-[11px] flex items-center justify-between text-[#C6A56B] font-ledger">
          <span>FILM / IP</span>
          <ArrowRight className="w-3.5 h-3.5 text-[#AAA5A0]" />
          <span>DISTRIBUTION</span>
          <ArrowRight className="w-3.5 h-3.5 text-[#AAA5A0]" />
          <span>REVENUE RECOGNITION</span>
          <ArrowRight className="w-3.5 h-3.5 text-[#AAA5A0]" />
          <span>COST MATCHING</span>
          <ArrowRight className="w-3.5 h-3.5 text-[#AAA5A0]" />
          <span className="text-[#829C80] font-bold">PROFIT</span>
        </div>
      </div>
    </div>
  );
}

// SLIDE 2: HARRY POTTER BOOK SALES (WITH VINTAGE LIBRARY IMAGE)
function Slide2Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  const [showExplanation, setShowExplanation] = useState(false);
  const data = [
    { year: 'FY2006', val: 195, label: 'US$ 195M', note: 'Half-Blood Prince impact' },
    { year: 'FY2007', val: 20, label: 'US$ 20M', note: 'Interim trough between major releases' },
    { year: 'FY2008', val: 270, label: 'US$ 270M', note: 'Deathly Hallows launch (Jul 2007)' },
  ];

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-center">
      {/* Left: Bar Chart & Library Image Companion */}
      <div className="lg:col-span-7 bg-[#101012]/70 border border-[#C6A56B]/25 rounded-lg p-4 flex flex-col justify-between h-full">
        <div>
          <div className="flex justify-between items-center">
            <span className="text-xs font-ledger text-[#C6A56B]">SCHOLASTIC CORP · REPORTED TRADE REVENUE</span>
            <button
              onClick={() =>
                onZoom(
                  ASSETS.vintageLibrary,
                  'Towering midnight library bookshelves and leatherbound financial ledgers: The publishing engine behind global literary phenomena.',
                  'SCHOLASTIC PUBLISHING ARCHIVE'
                )
              }
              className="flex items-center gap-1 text-[10px] text-[#C6A56B] hover:underline"
            >
              <Eye className="w-3 h-3" />
              <span>Library Archive</span>
            </button>
          </div>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] mt-1">Harry Potter Series Revenue Impact by Financial Year</h4>
        </div>

        {/* Visual Bar Graph */}
        <div className="flex items-end justify-around h-36 pt-2 pb-1 border-b border-[#C6A56B]/30">
          {data.map((item, i) => (
            <div key={i} className="flex flex-col items-center w-24 group">
              <span className="text-xs font-ledger text-[#C6A56B] font-bold mb-1">{item.label}</span>
              <div
                className="w-16 rounded-t transition-all duration-700 bg-gradient-to-t from-[#521F2B] to-[#C6A56B] group-hover:brightness-125 shadow-lg"
                style={{ height: `${(item.val / 270) * 105}px` }}
              />
              <span className="text-xs font-ledger text-[#F0E5D0] mt-1 font-semibold">{item.year}</span>
              <span className="text-[9px] text-[#AAA5A0] text-center leading-tight font-serif-body">{item.note}</span>
            </div>
          ))}
        </div>

        <div className="text-[10px] text-[#AAA5A0] flex justify-between pt-1">
          <span>Fiscal years end May 31. Reported in Scholastic Corp FY2008 Form 10-K.</span>
          <span className="text-[#C6A56B] font-ledger font-bold">+1,250% FY07→FY08 Rebound</span>
        </div>
      </div>

      {/* Right: Journal Entry Panel & Library Picture Card */}
      <div className="lg:col-span-5 flex flex-col justify-between h-full space-y-2.5">
        {/* Visual Card */}
        <div
          onClick={() =>
            onZoom(
              ASSETS.vintageLibrary,
              'Publishing ledger desk and archival hardcover volumes: Recording millions in book wholesale under ASC 606.',
              'TRADE PUBLISHING ACCOUNTING'
            )
          }
          className="h-20 w-full rounded-md overflow-hidden relative border border-[#C6A56B]/30 cursor-pointer group shadow-md"
        >
          <img
            src={ASSETS.vintageLibrary}
            alt="Scholastic Midnight Library Archive"
            referrerPolicy="no-referrer"
            className="w-full h-full object-cover group-hover:scale-105 transition duration-500 brightness-90"
          />
          <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-transparent to-transparent flex items-end p-2">
            <span className="text-[10px] text-[#F0E5D0] font-serif-body italic">
              Scholastic Children's Book Group · 11 Million copies shipped first 24h
            </span>
          </div>
        </div>

        <div className="bg-[#521F2B]/35 border border-[#C6A56B]/40 rounded-lg p-3">
          <div className="flex items-center justify-between text-xs font-ledger text-[#C6A56B] mb-1.5 font-bold uppercase">
            <span className="flex items-center gap-1.5">
              <FileSpreadsheet className="w-4 h-4" />
              Journal Entry · Book Trade Sales
            </span>
          </div>
          <div className="bg-[#101012] p-2.5 rounded font-ledger text-xs text-[#F0E5D0] space-y-1 border border-[#C6A56B]/20">
            <div className="flex justify-between">
              <span>Dr. Cash / Accounts Receivable</span>
              <span className="text-[#829C80] font-bold">$270,000,000</span>
            </div>
            <div className="flex justify-between pl-4 text-[#AAA5A0]">
              <span>Cr. Trade Sales Revenue</span>
              <span className="text-[#829C80] font-bold">$270,000,000</span>
            </div>
          </div>
          <p className="text-[10px] text-[#AAA5A0] mt-1.5 leading-relaxed font-serif-body">
            Revenue recognized under ASC 606 when control of printed books transfers to retailers (wholesale), net of return reserves.
          </p>
        </div>

        <div className="bg-[#101012]/80 border border-[#C6A56B]/20 rounded-lg p-2.5 text-[11px] space-y-1">
          <div className="flex justify-between">
            <span className="text-[#AAA5A0]">Gross Sales: Wholesale shipments</span>
            <span className="text-[#829C80] font-bold font-ledger">Publisher Income</span>
          </div>
        </div>
      </div>
    </div>
  );
}

// SLIDE 3: MERCHANDISE - MINISO & KINDER JOY
function Slide3Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-3.5 w-full h-full items-stretch">
      {/* Panel A: MINISO */}
      <div className="bg-[#101012]/75 border border-[#C6A56B]/20 rounded-lg p-3.5 flex flex-col justify-between">
        <div>
          <span className="text-[10px] font-ledger text-[#C6A56B] tracking-wider uppercase block mb-1">
            PANEL A · LIFESTYLE RETAIL
          </span>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] font-bold">MINISO × Harry Potter</h4>
          <p className="text-[11px] text-[#AAA5A0] mt-1.5 leading-relaxed font-serif-body">
            Global lifestyle brand collaboration featuring an officially licensed catalog of over{' '}
            <span className="text-[#C6A56B] font-bold">800+ SKUs</span> spanning stationery, bags, and plush collectibles.
          </p>
          <div className="mt-2.5 bg-[#521F2B]/40 p-2.5 rounded border border-[#C6A56B]/20 text-[10px] space-y-1">
            <div className="text-[#F0E5D0] font-semibold font-serif-heading">Reported Launch Performance:</div>
            <div className="text-[#C6A56B] font-ledger font-bold">Hong Kong pop-up: &gt;HK$ 5 Million in 7 days</div>
            <div className="text-[#AAA5A0] font-serif-body">MINISO recognizes gross product sales; licensor receives contractual royalty.</div>
          </div>
        </div>
        <div className="text-[9px] text-[#AAA5A0]/70 mt-2 font-ledger">Source: MINISO Brand Announcement (2024)</div>
      </div>

      {/* Panel B: Kinder Joy */}
      <div className="bg-[#101012]/75 border border-[#C6A56B]/20 rounded-lg p-3.5 flex flex-col justify-between">
        <div>
          <span className="text-[10px] font-ledger text-[#C6A56B] tracking-wider uppercase block mb-1">
            PANEL B · FMCG COLLECTIBLES
          </span>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] font-bold">Kinder Joy × Harry Potter</h4>
          <p className="text-[11px] text-[#AAA5A0] mt-1.5 leading-relaxed font-serif-body">
            Ferrero confectionary campaign combining chocolate eggs with miniature character figurines (Quidditch &amp; Hogwarts toys).
          </p>
          <div className="mt-2.5 bg-[#19171B] p-2.5 rounded border border-[#C6A56B]/20 text-[10px] space-y-1">
            <div className="text-[#F0E5D0] font-semibold font-serif-heading">Indian Consumer Phenomenon:</div>
            <div className="text-[#AAA5A0] font-serif-body">Widespread viral stockouts across Indian modern trade &amp; quick-commerce.</div>
            <div className="text-[#829C80] font-bold font-ledger">Accounting Caution:</div>
            <div className="text-[#AAA5A0] font-serif-body">Kinder Joy product revenues belong to Ferrero. Licensor records royalties only.</div>
          </div>
        </div>
        <div className="text-[9px] text-[#AAA5A0]/70 mt-2 font-ledger">Source: IIT Madras Article 19 (Feb 2025)</div>
      </div>

      {/* Panel C: Merchandise Accounting */}
      <div className="bg-[#521F2B]/35 border border-[#C6A56B]/40 rounded-lg p-3.5 flex flex-col justify-between">
        <div>
          <span className="text-[10px] font-ledger text-[#C6A56B] tracking-wider uppercase block mb-1">
            PANEL C · ACCOUNTING ENTRIES
          </span>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] font-bold">Inventory &amp; COGS Flow</h4>

          <div className="space-y-1.5 mt-2 font-ledger text-[10px]">
            <div className="bg-[#101012] p-2 rounded border border-[#C6A56B]/20">
              <span className="text-[#AAA5A0] block text-[9px]">1. On Purchase of Inventory</span>
              <span className="text-[#F0E5D0]">Dr. Inventory Asset</span>
              <br />
              <span className="text-[#AAA5A0] pl-2">Cr. Cash / Trade Payables</span>
            </div>

            <div className="bg-[#101012] p-2 rounded border border-[#C6A56B]/20">
              <span className="text-[#AAA5A0] block text-[9px]">2. On Retail Sale</span>
              <span className="text-[#F0E5D0]">Dr. Cash / Receivables</span>
              <br />
              <span className="text-[#AAA5A0] pl-2">Cr. Sales Revenue</span>
            </div>

            <div className="bg-[#101012] p-2 rounded border border-[#C6A56B]/20">
              <span className="text-[#AAA5A0] block text-[9px]">3. On Cost Recognition</span>
              <span className="text-[#F0E5D0]">Dr. Cost of Goods Sold (COGS)</span>
              <br />
              <span className="text-[#AAA5A0] pl-2">Cr. Inventory</span>
            </div>
          </div>
        </div>

        <div className="mt-1.5 pt-1.5 border-t border-[#C6A56B]/30 text-center font-ledger text-[10px] text-[#C6A56B] font-bold">
          GROSS PROFIT = SALES REVENUE − COGS
        </div>
      </div>
    </div>
  );
}

// SLIDE 4: ROYALTIES (SCHOLASTIC)
function Slide4Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-center">
      {/* Table Data Side */}
      <div className="lg:col-span-7 bg-[#101012]/70 border border-[#C6A56B]/30 rounded-lg p-4 flex flex-col justify-between h-full">
        <div>
          <span className="text-xs font-ledger text-[#C6A56B] uppercase">Scholastic Form 10-K Evidence</span>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] mt-1 mb-2">
            Impact of Franchise Release on Corporate Royalty Expenses (FY16 vs FY17)
          </h4>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs font-sans">
              <thead>
                <tr className="border-b border-[#C6A56B]/30 text-[#C6A56B] font-serif-heading">
                  <th className="py-1.5">Particulars (US$ Millions)</th>
                  <th className="py-1.5 text-right">FY2016</th>
                  <th className="py-1.5 text-right">FY2017</th>
                  <th className="py-1.5 text-right">Variance</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#AAA5A0]/10 text-[#F0E5D0]">
                <tr>
                  <td className="py-2 font-semibold text-[#F0E5D0]">Royalty Expense</td>
                  <td className="py-2 text-right font-ledger">$ 92.0 M</td>
                  <td className="py-2 text-right font-ledger text-[#C6A56B] font-bold">$ 146.4 M</td>
                  <td className="py-2 text-right font-ledger text-[#829C80] font-bold">+US$ 54.4 M (+59.1%)</td>
                </tr>
                <tr>
                  <td className="py-2 text-[#AAA5A0]">Total Cost of Goods Sold</td>
                  <td className="py-2 text-right font-ledger">$ 762.3 M</td>
                  <td className="py-2 text-right font-ledger">$ 814.5 M</td>
                  <td className="py-2 text-right font-ledger text-[#AAA5A0]">+US$ 52.2 M (+6.8%)</td>
                </tr>
                <tr>
                  <td className="py-2 text-[#AAA5A0]">Children's Publishing Op. Income</td>
                  <td className="py-2 text-right font-ledger">$ 120.6 M</td>
                  <td className="py-2 text-right font-ledger">$ 143.1 M</td>
                  <td className="py-2 text-right font-ledger text-[#829C80]">+US$ 22.5 M (+18.7%)</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <div className="bg-[#19171B] p-2.5 rounded border-l-2 border-[#C6A56B] text-[10px] text-[#AAA5A0] leading-relaxed font-serif-body">
          <span className="text-[#C6A56B] font-bold font-ledger">Filing Context:</span> The US$ 54.4M jump in royalty expenses coincided directly with the July 2016 publication of <span className="italic text-[#F0E5D0]">Harry Potter and the Cursed Child</span>, verifying how big IP drives variable licensing outlays.
        </div>
      </div>

      {/* Accounting Mechanism Side */}
      <div className="lg:col-span-5 flex flex-col justify-between h-full space-y-2.5">
        <div className="bg-[#521F2B]/40 border border-[#C6A56B]/40 rounded-lg p-3.5">
          <span className="text-xs font-ledger text-[#C6A56B] uppercase font-bold block mb-1.5">
            The Two Sides of a Royalty Contract
          </span>

          <div className="space-y-2 text-xs">
            <div className="bg-[#101012] p-2.5 rounded border border-[#C6A56B]/20">
              <span className="text-[#829C80] font-bold block text-[10px] mb-1 font-ledger">1. PUBLISHER (Licensee - Scholastic):</span>
              <div className="font-ledger text-[10px] text-[#F0E5D0]">
                Dr. Royalty Expense  $54.4M<br />
                <span className="pl-3 text-[#AAA5A0]">Cr. Royalty Payable  $54.4M</span>
              </div>
              <span className="text-[9px] text-[#AAA5A0] mt-0.5 block font-serif-body">Recorded as a direct Cost of Goods Sold.</span>
            </div>

            <div className="bg-[#101012] p-2.5 rounded border border-[#C6A56B]/20">
              <span className="text-[#C6A56B] font-bold block text-[10px] mb-1 font-ledger">2. RIGHTS HOLDER (Licensor - Author/Studio):</span>
              <div className="font-ledger text-[10px] text-[#F0E5D0]">
                Dr. Royalty Receivable<br />
                <span className="pl-3 text-[#AAA5A0]">Cr. Royalty Revenue</span>
              </div>
              <span className="text-[9px] text-[#AAA5A0] mt-0.5 block font-serif-body">Recognized under ASC 606 upon sales milestones.</span>
            </div>
          </div>
        </div>

        <div className="text-[10px] text-[#AAA5A0] bg-[#101012] p-2 rounded border border-[#C6A56B]/20 font-serif-body">
          *Figures reflect Scholastic company-wide reporting; do not assume the entire increase accrued to a single entity.
        </div>
      </div>
    </div>
  );
}

// SLIDE 5: FILM PRODUCTION COSTS & AMORTISATION (WBD)
function Slide5Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  const costCategories = [
    { label: 'Released, less amortisation', amount: 8692, pct: '59.3%' },
    { label: 'Completed, not yet released', amount: 1630, pct: '11.1%' },
    { label: 'In production and development', amount: 4326, pct: '29.6%' },
  ];

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-center">
      {/* Left Cost Distribution */}
      <div className="lg:col-span-6 bg-[#101012]/70 border border-[#C6A56B]/30 rounded-lg p-4 flex flex-col justify-between h-full">
        <div>
          <span className="text-xs font-ledger text-[#C6A56B] uppercase">Warner Bros. Discovery FY2025 10-K</span>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] mt-1 mb-1">Content Cost Capitalisation Breakdown</h4>
          <span className="text-xs text-[#829C80] font-ledger font-bold">Total Production Costs: US$ 14,648 Million</span>

          <div className="space-y-2.5 mt-3">
            {costCategories.map((cat, idx) => (
              <div key={idx} className="bg-[#19171B] p-2.5 rounded border border-[#C6A56B]/20">
                <div className="flex justify-between text-xs mb-1">
                  <span className="text-[#F0E5D0]">{cat.label}</span>
                  <span className="font-ledger text-[#C6A56B] font-bold">US$ {cat.amount}M ({cat.pct})</span>
                </div>
                <div className="w-full bg-[#101012] h-2 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-[#521F2B] to-[#C6A56B]"
                    style={{ width: cat.pct }}
                  />
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="text-[10px] text-[#AAA5A0] mt-2 pt-2 border-t border-[#C6A56B]/20 font-serif-body">
          Films are not expensed immediately. Spending is capitalised as content assets and amortised using the individual-film-forecast computation method (ASC 926).
        </div>
      </div>

      {/* Right Accounting Pathway */}
      <div className="lg:col-span-6 flex flex-col justify-between h-full space-y-2.5">
        <div className="bg-[#521F2B]/35 border border-[#C6A56B]/30 rounded-lg p-3.5">
          <span className="text-xs font-serif-heading uppercase tracking-wider text-[#C6A56B] font-bold block mb-1.5">
            The Content Accounting Lifecycle
          </span>

          <div className="space-y-2 text-xs">
            <div className="flex items-start gap-2">
              <span className="w-5 h-5 rounded-full bg-[#C6A56B] text-[#101012] font-ledger font-bold flex items-center justify-center text-[10px] shrink-0 mt-0.5">1</span>
              <div>
                <span className="text-[#F0E5D0] font-semibold">Production Spending:</span>
                <p className="text-[10px] text-[#AAA5A0] font-serif-body">Cash spent on sets, CGI, cast salaries is capitalised to Balance Sheet as Film Inventory Asset.</p>
              </div>
            </div>

            <div className="flex items-start gap-2">
              <span className="w-5 h-5 rounded-full bg-[#C6A56B] text-[#101012] font-ledger font-bold flex items-center justify-center text-[10px] shrink-0 mt-0.5">2</span>
              <div>
                <span className="text-[#F0E5D0] font-semibold">Release &amp; Availability:</span>
                <p className="text-[10px] text-[#AAA5A0] font-serif-body">Asset shifts from 'In Production' to 'Released Content, less amortisation'.</p>
              </div>
            </div>

            <div className="flex items-start gap-2">
              <span className="w-5 h-5 rounded-full bg-[#829C80] text-[#101012] font-ledger font-bold flex items-center justify-center text-[10px] shrink-0 mt-0.5">3</span>
              <div>
                <span className="text-[#F0E5D0] font-semibold">Amortisation:</span>
                <p className="text-[10px] text-[#AAA5A0] font-serif-body">Expense recognized over time matching ultimate projected revenues across windows.</p>
              </div>
            </div>
          </div>
        </div>

        {/* Illustrative Entry */}
        <div className="bg-[#101012] p-2.5 rounded-lg border border-[#C6A56B]/30 font-ledger text-xs">
          <span className="text-[9px] text-[#C6A56B] uppercase block mb-1">Illustrative Journal Entry · Periodic Amortisation</span>
          <div className="text-[#F0E5D0]">Dr. Content Amortisation Expense</div>
          <div className="text-[#AAA5A0] pl-4">Cr. Film &amp; Television Asset (Balance Sheet)</div>
        </div>
      </div>
    </div>
  );
}

// SLIDE 6: INDIAN CINEMA - VALUE CHAIN & PVR INOX (WITH MULTIPLEX IMAGE)
function Slide6Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  const exhibitorStreams = [
    { name: 'Net Ticket Admissions', share: '50-55%', margin: 'Medium', desc: 'Share of box office after GST (12/18%) & distributor splits' },
    { name: 'Food & Beverage (F&B)', share: '28-32%', margin: '70-75% Gross', desc: 'Popcorn, beverages, hot foods with minimal distributor encumbrance' },
    { name: 'Screen Advertising', share: '8-10%', margin: '>80% EBITDA', desc: 'Commercials and brand integrations prior to showtime' },
    { name: 'Convenience Fees & Other', share: '5-7%', margin: 'High', desc: 'Digital booking fees, venue rentals, and arcade revenues' },
  ];

  return (
    <div className="flex flex-col justify-between w-full h-full space-y-2.5">
      {/* Value Chain Top Banner with Multiplex Image Button */}
      <div className="bg-[#521F2B]/40 border border-[#C6A56B]/30 p-2 rounded-lg flex items-center justify-between text-xs text-[#F0E5D0]">
        <div className="flex items-center gap-1.5">
          <span className="font-ledger text-[#C6A56B] font-bold">PRODUCER</span>
          <span className="text-[10px] text-[#AAA5A0]">(Creates IP)</span>
        </div>
        <ArrowRight className="w-3.5 h-3.5 text-[#C6A56B]" />
        <div className="flex items-center gap-1.5">
          <span className="font-ledger text-[#C6A56B] font-bold">DISTRIBUTOR</span>
          <span className="text-[10px] text-[#AAA5A0]">(Rights)</span>
        </div>
        <ArrowRight className="w-3.5 h-3.5 text-[#C6A56B]" />
        <div className="flex items-center gap-1.5">
          <span className="font-ledger text-[#C6A56B] font-bold">EXHIBITOR (PVR INOX)</span>
          <span className="text-[10px] text-[#829C80] font-bold">(Screens &amp; Concessions)</span>
        </div>
        <ArrowRight className="w-3.5 h-3.5 text-[#C6A56B]" />
        <div className="flex items-center gap-1.5">
          <span className="font-ledger text-[#C6A56B] font-bold">PATRON</span>
          <span className="text-[10px] text-[#AAA5A0]">(F&amp;B + Tickets)</span>
        </div>
        <button
          onClick={() =>
            onZoom(
              ASSETS.indianMultiplex,
              'Grand entrance of Rajmahal Luxury Multiplex and Gourmet Confectionery Lounge: The high-margin F&B engine that powers Indian exhibition.',
              'PVR INOX EXHIBITION FOYER'
            )
          }
          className="ml-2 px-2 py-0.5 rounded bg-[#C6A56B]/20 text-[#C6A56B] text-[10px] font-ledger flex items-center gap-1 hover:bg-[#C6A56B]/40 transition"
        >
          <Eye className="w-3 h-3" /> Foyer Photo
        </button>
      </div>

      {/* Exhibitor Revenue Breakdown with Visual Inset */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-2.5 flex-1">
        {exhibitorStreams.map((st, i) => (
          <div key={i} className="bg-[#101012]/75 border border-[#C6A56B]/20 rounded-lg p-2.5 flex flex-col justify-between">
            <div>
              <span className="text-[10px] font-ledger text-[#C6A56B] font-bold">STREAM 0{i + 1}</span>
              <h5 className="text-xs font-serif-heading text-[#F0E5D0] font-bold mt-0.5">{st.name}</h5>
              <div className="mt-1.5 space-y-0.5 text-[10px]">
                <div className="flex justify-between">
                  <span className="text-[#AAA5A0]">Revenue Mix:</span>
                  <span className="font-ledger text-[#F0E5D0]">{st.share}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-[#AAA5A0]">Margin:</span>
                  <span className="font-ledger text-[#829C80] font-semibold">{st.margin}</span>
                </div>
              </div>
              <p className="text-[9px] text-[#AAA5A0] mt-1.5 leading-relaxed font-serif-body">{st.desc}</p>
            </div>
            <div className="pt-1.5 border-t border-[#C6A56B]/10 text-[9px] text-[#C6A56B] font-ledger">PVR INOX Disclosures</div>
          </div>
        ))}
      </div>

      <div className="text-[10px] text-[#AAA5A0] bg-[#19171B] p-2 rounded border-l-4 border-[#C6A56B] flex justify-between items-center font-serif-body">
        <span>
          <strong className="text-[#F0E5D0]">Accounting Note:</strong> Gross Box Office Collections are subject to GST before distributor split. F&amp;B sales carry no distributor sharing and represent the primary operational profit driver.
        </span>
        <span className="font-ledger text-[#C6A56B] hidden sm:inline">Ind AS 115 Compliance</span>
      </div>
    </div>
  );
}

// SLIDE 7: PVR INOX BALANCE SHEET (WITH MULTIPLEX LEASE ASSET PHOTO)
function Slide7Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-center">
      {/* Audited Balance Sheet Extract */}
      <div className="lg:col-span-7 bg-[#101012]/75 border border-[#C6A56B]/30 rounded-lg p-3.5 flex flex-col justify-between h-full">
        <div>
          <div className="flex justify-between items-center mb-1.5">
            <span className="text-xs font-ledger text-[#C6A56B] uppercase font-bold">PVR INOX LIMITED · CONSOLIDATED</span>
            <span className="text-[10px] text-[#AAA5A0] font-ledger">As on 31 March 2025</span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs font-sans">
              <thead>
                <tr className="border-b border-[#C6A56B]/40 text-[#C6A56B] font-serif-heading">
                  <th className="py-1">Balance Sheet Particulars</th>
                  <th className="py-1 text-right">Amount (₹ Millions)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#AAA5A0]/10 text-[#F0E5D0] font-ledger text-[11px]">
                <tr>
                  <td className="py-1 font-sans">Equity Share Capital</td>
                  <td className="py-1 text-right">₹ 982 M</td>
                </tr>
                <tr>
                  <td className="py-1 font-sans text-[#AAA5A0]">Other Equity (Reserves &amp; Surplus)</td>
                  <td className="py-1 text-right">₹ 69,265 M</td>
                </tr>
                <tr className="bg-[#521F2B]/20 text-[#C6A56B] font-bold">
                  <td className="py-1 font-sans">Total Equity (Shareholders' Funds)</td>
                  <td className="py-1 text-right">₹ 70,247 M</td>
                </tr>
                <tr>
                  <td className="py-1 font-sans">Long-term Borrowings</td>
                  <td className="py-1 text-right">₹ 9,198 M</td>
                </tr>
                <tr>
                  <td className="py-1 font-sans">Current Liabilities</td>
                  <td className="py-1 text-right">₹ 24,977 M</td>
                </tr>
                <tr>
                  <td className="py-1 font-sans text-[#829C80]">Non-Current Assets (inc. Right-of-Use Leases)</td>
                  <td className="py-1 text-right text-[#829C80]">₹ 145,920 M</td>
                </tr>
                <tr>
                  <td className="py-1 font-sans">Current Assets</td>
                  <td className="py-1 text-right">₹ 10,810 M</td>
                </tr>
                <tr className="border-t-2 border-[#C6A56B] text-[#F0E5D0] font-bold bg-[#101012]">
                  <td className="py-1.5 font-sans">Total Assets = Equity + Liabilities</td>
                  <td className="py-1.5 text-right text-[#C6A56B]">₹ 156,730 M</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <div className="text-[10px] text-[#AAA5A0] pt-1.5 border-t border-[#C6A56B]/20 flex justify-between font-ledger">
          <span>Audited Indian Accounting Standards (Ind AS) Filing</span>
          <span className="text-[#829C80]">Balanced: Assets = Equity + Liab</span>
        </div>
      </div>

      {/* Financial Analytical Takeaways & Multiplex Asset Photo */}
      <div className="lg:col-span-5 flex flex-col justify-between h-full space-y-2">
        {/* Multiplex Photo Inset */}
        <div
          onClick={() =>
            onZoom(
              ASSETS.indianMultiplex,
              'Multiplex Real Estate & Screens (₹145,920 Million Non-Current Assets under Ind AS 116 Lease Capitalisation).',
              'TANGIBLE MULTIPLEX ASSET BASE'
            )
          }
          className="h-20 w-full rounded-md overflow-hidden relative border border-[#C6A56B]/30 cursor-pointer group shadow-md"
        >
          <img
            src={ASSETS.indianMultiplex}
            alt="PVR INOX Multiplex Infrastructure"
            referrerPolicy="no-referrer"
            className="w-full h-full object-cover group-hover:scale-105 transition duration-500 brightness-95"
          />
          <div className="absolute inset-0 bg-gradient-to-t from-black/85 via-transparent to-transparent flex items-end p-2">
            <span className="text-[10px] text-[#F0E5D0] font-serif-body italic">
              ₹145.9 Billion Non-Current Assets: Screens, Projectors &amp; Leases
            </span>
          </div>
        </div>

        <div className="bg-[#521F2B]/35 border border-[#C6A56B]/30 rounded-lg p-3">
          <span className="text-xs font-ledger text-[#C6A56B] font-bold uppercase block mb-1">
            Accounting Analysis &amp; Ind AS 116 Impact
          </span>
          <div className="space-y-1 text-[10.5px] text-[#AAA5A0] font-serif-body">
            <div>
              <span className="text-[#F0E5D0] font-semibold block">1. Asset-Heavy Multiplex Footprint:</span>
              <span>Non-current assets reflect 4K laser projectors, Dolby Atmos, and massive leasehold capitalisations.</span>
            </div>
            <div>
              <span className="text-[#F0E5D0] font-semibold block">2. Balance Sheet ≠ Profit:</span>
              <span>An enterprise holding ₹156,730M in assets may still record quarterly accounting net losses due to heavy depreciation.</span>
            </div>
          </div>
        </div>

        <div className="p-2 bg-[#101012] border border-[#C6A56B]/30 rounded-lg text-center font-ledger text-[11px] text-[#C6A56B]">
          ASSETS (₹156,730M) = EQUITY (₹70,247M) + LIABILITIES (₹86,483M)
        </div>
      </div>
    </div>
  );
}

// SLIDE 8: DHURANDHAR - BOX OFFICE VS PRODUCER PROFIT (WITH BALANCE SHEET LEDGER PHOTO)
function Slide8Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-center">
      {/* Waterfall Concept Box */}
      <div className="lg:col-span-7 bg-[#101012]/70 border border-[#C6A56B]/30 rounded-lg p-4 flex flex-col justify-between h-full">
        <div>
          <div className="flex justify-between items-center">
            <span className="text-xs font-ledger text-[#C6A56B] uppercase">Industry Financial Breakdown · Feb 2026</span>
            <button
              onClick={() =>
                onZoom(
                  ASSETS.boxOfficeWaterfall,
                  'Golden Clapperboard, fountain pen and confidential audited balance sheet: The rigorous math behind theatrical waterfalls.',
                  'CINEMATIC FINANCIAL AUDIT'
                )
              }
              className="flex items-center gap-1 text-[10px] text-[#C6A56B] hover:underline"
            >
              <Eye className="w-3 h-3" />
              <span>Ledger Photo</span>
            </button>
          </div>
          <h4 className="text-base font-serif-heading text-[#F0E5D0] mt-1 mb-2">Dhurandhar Box-Office Waterfall Analysis</h4>

          {/* Stepped Waterfall */}
          <div className="space-y-1.5 text-xs">
            <div className="bg-[#19171B] p-2 rounded border border-[#C6A56B]/40 flex justify-between items-center">
              <div>
                <span className="text-[#C6A56B] font-bold">1. Worldwide Gross Box Office:</span>
                <span className="text-[10px] text-[#AAA5A0] block font-serif-body">Aggregated ticket admissions worldwide</span>
              </div>
              <span className="font-ledger text-base font-bold text-[#F0E5D0]">~₹ 1,250 Cr</span>
            </div>

            <div className="bg-[#19171B] p-2 rounded border border-[#C6A56B]/20 flex justify-between items-center">
              <div>
                <span className="text-[#F0E5D0] font-semibold">2. India Gross Collections:</span>
                <span className="text-[10px] text-[#AAA5A0] block font-serif-body">Domestic theatrical gross across all states</span>
              </div>
              <span className="font-ledger text-sm font-semibold text-[#F0E5D0]">~₹ 950 Cr</span>
            </div>

            <div className="bg-[#521F2B]/40 p-2 rounded border border-[#C6A56B]/30 flex justify-between items-center">
              <div>
                <span className="text-[#C6A56B] font-semibold">Less: Estimated India Theatre Share:</span>
                <span className="text-[10px] text-[#AAA5A0] block font-serif-body">Retained by multiplex and single-screen owners</span>
              </div>
              <span className="font-ledger text-sm font-semibold text-red-300">- ~₹ 300 Cr</span>
            </div>

            <div className="bg-[#101012] p-2 rounded border-l-4 border-[#829C80] flex justify-between items-center">
              <div>
                <span className="text-[#829C80] font-bold">Estimated Producer Theatrical Receipts:</span>
                <span className="text-[10px] text-[#AAA5A0] block font-serif-body">Net theatrical proceeds accruing to Jio &amp; B62 Studios</span>
              </div>
              <span className="font-ledger text-sm font-bold text-[#829C80]">Subject to contract</span>
            </div>
          </div>
        </div>

        <div className="text-[10px] text-[#AAA5A0] pt-1.5 border-t border-[#C6A56B]/20 font-ledger">
          Source: The Week Box-Office Financial Analysis (Published 03 Feb 2026).
        </div>
      </div>

      {/* Accounting Formula & Reality Panel with Image Companion */}
      <div className="lg:col-span-5 flex flex-col justify-between h-full space-y-2.5">
        {/* Ledger Still Life Card */}
        <div
          onClick={() =>
            onZoom(
              ASSETS.boxOfficeWaterfall,
              'Confidential Balance Sheet ledger, golden clapperboard and fountain pen: Why viral collections do not equal producer net income.',
              'AUDITED PRODUCER WATERFALL'
            )
          }
          className="h-20 w-full rounded-md overflow-hidden relative border border-[#C6A56B]/30 cursor-pointer group shadow-md"
        >
          <img
            src={ASSETS.boxOfficeWaterfall}
            alt="Audited Balance Sheet Ledger Still Life"
            referrerPolicy="no-referrer"
            className="w-full h-full object-cover group-hover:scale-105 transition duration-500 brightness-95"
          />
          <div className="absolute inset-0 bg-gradient-to-t from-black/85 via-transparent to-transparent flex items-end p-2">
            <span className="text-[10px] text-[#F0E5D0] font-serif-body italic">
              Production Ledger · Scene 42 Take 3
            </span>
          </div>
        </div>

        <div className="bg-[#521F2B]/35 border border-[#C6A56B]/40 rounded-lg p-3">
          <span className="text-xs font-ledger text-[#C6A56B] font-bold uppercase block mb-1">
            The Producer Net Profit Formula
          </span>

          <div className="bg-[#101012] p-2 rounded font-ledger text-[10px] text-[#C6A56B] leading-relaxed border border-[#C6A56B]/20 mb-2">
            PRODUCER PROFIT = RECOGNISED PRODUCER REVENUE<br />
            − PRODUCTION COSTS − P&amp;A − TALENT SHARES
          </div>

          <div className="bg-[#101012]/90 p-2 rounded border-l-2 border-red-500 text-[10.5px] text-[#AAA5A0] leading-relaxed font-serif-body">
            <strong className="text-[#F0E5D0] block mb-0.5">Auditing Reality Check:</strong>
            "Exact producer profit cannot be reliably calculated from gross box-office collections alone. Private production budgets and distribution agreements are not public statutory filings."
          </div>
        </div>

        <div className="text-[10px] text-[#AAA5A0] bg-[#101012] p-2 rounded border border-[#C6A56B]/20 font-serif-body">
          Digital, satellite, and music rights often contribute more to net producer recovery than theatrical splits in modern Indian cinema.
        </div>
      </div>
    </div>
  );
}

// SLIDE 9: HOLLYWOOD VS INDIAN CINEMA COMPARISON
function Slide9Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  const comparison = [
    { dim: 'Primary Revenue Engine', hw: 'Content licensing, SVOD, gaming, retail IP', ind: 'Ticket admissions, F&B concessions, screen ads' },
    { dim: 'Licensing & Merchandise', hw: 'Primary margin driver (MINISO, Kinder Joy)', ind: 'Incidental / nascent retail penetration' },
    { dim: 'Production Costs', hw: 'Capitalised content library ($14.6B at WBD)', ind: 'Exhibitor incurs film hire; does not produce' },
    { dim: 'Inventory Accounting', hw: 'Wholesale retail goods & home media stock', ind: 'Perishable F&B goods (popcorn, syrups)' },
    { dim: 'Balance Sheet Character', hw: 'Intangible IP assets & content libraries', ind: 'Tangible screens, cinema leases, projectors' },
    { dim: 'Core Financial Target', hw: 'Maximise lifetime global IP monetisation', ind: 'Maximise spend per head (SPH) & seat fill' },
  ];

  return (
    <div className="flex flex-col justify-between w-full h-full space-y-2.5">
      {/* Side-by-side Matrix */}
      <div className="overflow-x-auto bg-[#101012]/70 border border-[#C6A56B]/30 rounded-lg p-2.5">
        <table className="w-full text-left text-xs font-sans">
          <thead>
            <tr className="border-b border-[#C6A56B]/40 font-serif-heading">
              <th className="py-1.5 text-[#C6A56B] w-1/4">Accounting Dimension</th>
              <th className="py-1.5 text-[#F0E5D0] w-3/8">Hollywood Studio (Warner Bros. Discovery)</th>
              <th className="py-1.5 text-[#C6A56B] w-3/8">Indian Cinema Exhibitor (PVR INOX)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#AAA5A0]/10 text-[11px]">
            {comparison.map((row, idx) => (
              <tr key={idx} className="hover:bg-[#19171B]/50 transition">
                <td className="py-1.5 font-semibold text-[#C6A56B] font-ledger text-[10px]">{row.dim}</td>
                <td className="py-1.5 text-[#F0E5D0] font-serif-body">{row.hw}</td>
                <td className="py-1.5 text-[#AAA5A0] font-serif-body">{row.ind}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Visual Models Summary with Dual Image Inset */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        <div
          onClick={() =>
            onZoom(
              ASSETS.hollywoodStudio,
              'Hollywood Studio Model: Driven by content asset creation, global syndication, and multi-decade IP exploitation.',
              'HOLLYWOOD STUDIO PARADIGM'
            )
          }
          className="bg-[#521F2B]/35 border border-[#C6A56B]/30 p-2 rounded text-xs flex justify-between items-center cursor-pointer hover:border-[#C6A56B] transition"
        >
          <div className="flex items-center gap-2">
            <img
              src={ASSETS.hollywoodStudio}
              alt="Hollywood Studio"
              referrerPolicy="no-referrer"
              className="w-8 h-8 rounded object-cover border border-[#C6A56B]/40"
            />
            <span className="font-serif-heading font-bold text-[#F0E5D0]">HOLLYWOOD STUDIO:</span>
          </div>
          <span className="font-ledger text-[10px] text-[#C6A56B]">Content + Licensing + Games</span>
        </div>

        <div
          onClick={() =>
            onZoom(
              ASSETS.indianMultiplex,
              'Indian Exhibitor Model: Driven by theater seat fill, concession spends, and Ind AS 116 leasehold commitments.',
              'INDIAN EXHIBITOR PARADIGM'
            )
          }
          className="bg-[#19171B] border border-[#C6A56B]/30 p-2 rounded text-xs flex justify-between items-center cursor-pointer hover:border-[#C6A56B] transition"
        >
          <div className="flex items-center gap-2">
            <img
              src={ASSETS.indianMultiplex}
              alt="Indian Cinema"
              referrerPolicy="no-referrer"
              className="w-8 h-8 rounded object-cover border border-[#C6A56B]/40"
            />
            <span className="font-serif-heading font-bold text-[#F0E5D0]">INDIAN EXHIBITOR:</span>
          </div>
          <span className="font-ledger text-[10px] text-[#829C80]">Tickets + F&amp;B + Screen Ads</span>
        </div>
      </div>

      <div className="text-[10px] text-[#AAA5A0] text-center font-serif-body">
        *Important: Studio and exhibitor financials are not directly comparable without accounting for differing business models, currencies, and reporting frameworks.
      </div>
    </div>
  );
}

// SLIDE 10: TAKEAWAYS & PRIMARY SOURCE ARCHIVE (WITH WATERFALL STILL LIFE)
function Slide10Content({ onZoom }: { onZoom: (src: string, caption: string, credit: string) => void }) {
  const sources = [
    { name: 'Scholastic FY2008 Form 10-K', link: 'https://www.sec.gov/Archives/edgar/data/866729/000093041308004578/c54272_10-k.htm', tag: 'US SEC' },
    { name: 'Scholastic FY2017 Form 10-K', link: 'https://www.sec.gov/Archives/edgar/data/866729/000086672917000009/schl5312017-10k.htm', tag: 'US SEC' },
    { name: 'Warner Bros. Discovery FY2025 Form 10-K', link: 'https://www.sec.gov/Archives/edgar/data/1437107/000143710726000020/wbd-20251231.htm', tag: 'US SEC' },
    { name: 'PVR INOX FY2024-25 Annual Report', link: 'https://nsearchives.nseindia.com/corporate/PVR_14082025115149_PVRINOXAGMNoticeAnnualReportFY20242514082025.pdf', tag: 'NSE India' },
    { name: 'WIPO — Harry Potter and the IP Bonanza', link: 'https://www.wipo.int/en/web/wipo-magazine/articles/harry-potter-and-the-ip-bonanza-35915', tag: 'WIPO' },
    { name: 'MINISO Global Collection Announcement', link: 'https://www.miniso.com/brand/news_56.html', tag: 'Corporate' },
    { name: 'Kinder Joy × Harry Potter Analysis', link: 'https://articlenineteeniitm.com/2025/02/18/harry-potter-kinder-joy-and-nostalgia-for-sale/', tag: 'IIT Madras' },
    { name: 'The Week — Dhurandhar Box Office Audit', link: 'https://www.theweek.in/news/entertainment/2026/02/03/dhurandhar-final-box-office-collection-analysis-how-much-money-did-jio-studios-and-b62-studios-make-from-theatres-in-india.html', tag: 'Editorial' },
    { name: 'Ministry of Corporate Affairs — Ind AS 115', link: 'https://www.mca.gov.in/Ministry/pdf/INDAS115.pdf', tag: 'Statutory' },
  ];

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 w-full h-full items-stretch">
      {/* 3 Core Rules with Final Still Life Banner */}
      <div className="lg:col-span-6 bg-[#521F2B]/40 border border-[#C6A56B]/30 rounded-lg p-4 flex flex-col justify-between">
        <div>
          <div className="flex justify-between items-center mb-1">
            <span className="text-xs font-ledger text-[#C6A56B] uppercase font-bold tracking-wider block">
              The Three Principles of Entertainment Accounting
            </span>
            <button
              onClick={() =>
                onZoom(
                  ASSETS.boxOfficeWaterfall,
                  'Final Audit Ledger Still Life: Grounded in US GAAP ASC 606 & Indian Accounting Standards (Ind AS).',
                  'ACCOUNTANCY FOR ENGINEERS LEDGER'
                )
              }
              className="flex items-center gap-1 text-[10px] text-[#C6A56B] hover:underline"
            >
              <Eye className="w-3 h-3" />
              <span>Full Ledger</span>
            </button>
          </div>

          <div className="space-y-2.5 text-xs mt-2">
            <div className="bg-[#101012] p-2.5 rounded border-l-2 border-[#C6A56B]">
              <span className="text-[#C6A56B] font-bold block mb-0.5 font-ledger text-[11px]">1. Analyse Revenue by Stream:</span>
              <p className="text-[11px] text-[#AAA5A0] font-serif-body">Never conflate gross theatrical collections with recognized studio revenue or exhibitor margins.</p>
            </div>

            <div className="bg-[#101012] p-2.5 rounded border-l-2 border-[#829C80]">
              <span className="text-[#829C80] font-bold block mb-0.5 font-ledger text-[11px]">2. Track Costs Correctly:</span>
              <p className="text-[11px] text-[#AAA5A0] font-serif-body">Distinguish capitalised content assets and inventory from direct royalty and operational costs.</p>
            </div>

            <div className="bg-[#101012] p-2.5 rounded border-l-2 border-[#C6A56B]">
              <span className="text-[#C6A56B] font-bold block mb-0.5 font-ledger text-[11px]">3. Evaluate the Complete Financial Position:</span>
              <p className="text-[11px] text-[#AAA5A0] font-serif-body">Box office headlines do not reflect balance-sheet lease liabilities, debt service, or true net profitability.</p>
            </div>
          </div>
        </div>

        <div className="mt-2.5 pt-2 border-t border-[#C6A56B]/20 text-[10.5px] italic font-serif-body text-[#F0E5D0]">
          "Entertainment accounting is not just about how much a film earns. It is about how rights are monetised, costs are recognised, assets are measured and financial performance is evaluated."
        </div>
      </div>

      {/* Complete Verified Statutory Archive */}
      <div className="lg:col-span-6 bg-[#101012]/80 border border-[#C6A56B]/30 rounded-lg p-4 flex flex-col justify-between">
        <div>
          <span className="text-xs font-ledger text-[#C6A56B] uppercase font-bold tracking-wider block mb-1.5">
            Verified Regulatory &amp; Research Ledger
          </span>

          <div className="space-y-1 mt-1">
            {sources.map((src, i) => (
              <a
                key={i}
                href={src.link}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center justify-between p-1.5 rounded bg-[#19171B] hover:bg-[#521F2B]/30 border border-[#C6A56B]/10 hover:border-[#C6A56B]/50 transition text-[11px]"
              >
                <span className="text-[#F0E5D0] font-serif-body truncate max-w-xs">{src.name}</span>
                <div className="flex items-center gap-1.5 text-[9px] text-[#C6A56B] font-ledger">
                  <span>{src.tag}</span>
                  <ExternalLink className="w-2.5 h-2.5" />
                </div>
              </a>
            ))}
          </div>
        </div>

        <div className="text-[9px] text-[#AAA5A0] pt-1.5 border-t border-[#C6A56B]/20 text-center font-serif-body">
          Case study prepared for university-level Accountancy for Engineers · Full Compliance with US GAAP &amp; Ind AS.
        </div>
      </div>
    </div>
  );
}
