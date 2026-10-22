import { SLIDES_META } from '../data/slidesData';

// Color definitions for Google Slides API (RGB 0..1 scale)
const COLOR_OBSIDIAN = { red: 0.063, green: 0.063, blue: 0.071 }; // #101012
const COLOR_GOLD = { red: 0.776, green: 0.647, blue: 0.420 }; // #C6A56B
const COLOR_PARCHMENT = { red: 0.941, green: 0.898, blue: 0.816 }; // #F0E5D0
const COLOR_MUTED = { red: 0.667, green: 0.647, blue: 0.627 }; // #AAA5A0
const COLOR_BURGUNDY = { red: 0.322, green: 0.122, blue: 0.169 }; // #521F2B
const COLOR_MIDNIGHT = { red: 0.098, green: 0.090, blue: 0.106 }; // #19171B
const COLOR_GREEN = { red: 0.510, green: 0.612, blue: 0.502 }; // #829C80

interface SlideContentSummary {
  leftBoxTitle: string;
  leftBoxText: string;
  rightBoxTitle: string;
  rightBoxText: string;
}

const SLIDES_CONTENT: Record<number, SlideContentSummary> = {
  1: {
    leftBoxTitle: "WARNER BROS. DISCOVERY (FY2025)\nAudited Company-Wide Figures",
    leftBoxText:
      "• Total Company Revenue: US$ 37,296 Million\n• Studios Segment Revenue: US$ 12,619 Million\n\n*Note: Consolidated WBD filings represent multi-franchise totals and corporate operations, not Harry Potter-only revenue.",
    rightBoxTitle: "THE 6 REVENUE PATHWAYS OF CINEMATIC IP",
    rightBoxText:
      "1. Theatrical Box Office: First window theatrical exhibition receipts\n2. Streaming & Subscriptions: SVOD platforms (Max global subscribers)\n3. Content Licensing: Television networks, third-party syndication\n4. Merchandise & Brand Licensing: Apparel, toys, retail lines\n5. Video Games & Interactive: Hogwarts Legacy ($1B+ retail sales)\n6. Home Entertainment: EST, VOD downloads, physical disc media",
  },
  2: {
    leftBoxTitle: "SCHOLASTIC CORP (FY06–FY08)\nReported Trade Revenue",
    leftBoxText:
      "• FY2006: US$ 195.0 Million (Post Half-Blood Prince)\n• FY2007: US$ 20.0 Million (Interim trough)\n• FY2008: US$ 270.0 Million (Deathly Hallows release)\n\nVariance: +1,250% rebound between FY07 and FY08 demonstrates extreme release-driven revenue elasticity.",
    rightBoxTitle: "ILLUSTRATIVE JOURNAL ENTRY (ASC 606)",
    rightBoxText:
      "Dr. Cash / Accounts Receivable    $270,000,000\n    Cr. Trade Sales Revenue           $270,000,000\n\n[Recording wholesale book revenue upon transfer of control to retailers, net of estimated return provisions.]\n\nDirect Manufacturing COGS & Author Royalties are recognized simultaneously to achieve revenue-expense matching.",
  },
  3: {
    leftBoxTitle: "PANEL A: MINISO RETAIL & PANEL B: KINDER JOY",
    leftBoxText:
      "MINISO COLLABORATION:\n• 800+ SKU lifestyle product collection worldwide\n• Pop-up generated >HK$ 5M in first 7 days\n• MINISO books gross product sales; pays royalty to licensor.\n\nFERRERO KINDER JOY TIE-IN:\n• Character mini-figurines inside chocolate eggs\n• High-velocity FMCG run across modern trade & q-commerce\n• Ferrero books retail sales; licensor records contractual royalty.",
    rightBoxTitle: "PANEL C: MERCHANDISE ACCOUNTING FLOW",
    rightBoxText:
      "1. Inventory Acquisition:\n   Dr. Inventory | Cr. Trade Payables / Cash\n\n2. Point of Sale:\n   Dr. Cash or Accounts Receivable | Cr. Sales Revenue\n\n3. Cost Recognition:\n   Dr. Cost of Goods Sold (COGS) | Cr. Inventory\n\nKEY METRIC: Gross Profit = Sales Revenue − COGS",
  },
  4: {
    leftBoxTitle: "SCHOLASTIC FORM 10-K EVIDENCE (FY16 vs FY17)",
    leftBoxText:
      "• Royalty Expense: $92.0M (FY16) → $146.4M (FY17)\n  [+$54.4M (+59.1%) increase]\n• Total Cost of Goods Sold: $762.3M → $814.5M (+6.8%)\n• Children's Publishing Op. Inc: $120.6M → $143.1M (+18.7%)\n\nFiling Context: Coincided directly with the July 2016 release of Harry Potter and the Cursed Child.",
    rightBoxTitle: "THE TWO SIDES OF A ROYALTY CONTRACT",
    rightBoxText:
      "LICENSEE (Scholastic / Publisher):\nDr. Royalty Expense    $54,400,000\n    Cr. Royalty Payable    $54,400,000\n(Recorded as a direct Cost of Goods Sold component)\n\nLICENSOR (Author / Studio):\nDr. Royalty Receivable\n    Cr. Royalty Revenue\n(Recognized under ASC 606 upon fulfillment of contract obligations)",
  },
  5: {
    leftBoxTitle: "WARNER BROS. DISCOVERY CONTENT ASSET BASE",
    leftBoxText:
      "Capitalised Content Costs Breakdown (FY2025 Form 10-K):\n• Released production costs (less amort.): $8,692M (59.3%)\n• Completed, not yet released: $1,630M (11.1%)\n• In production and development: $4,326M (29.6%)\n• Total Content Asset Value: US$ 14,648 Million",
    rightBoxTitle: "CONTENT ACCOUNTING LIFECYCLE (US GAAP ASC 926)",
    rightBoxText:
      "1. Production Outlays:\n   Production costs are capitalised as balance sheet assets, NOT expensed immediately upon cash payment.\n\n2. Release & Availability:\n   Transferred to released inventory upon public exhibition.\n\n3. Periodic Amortisation:\n   Dr. Content Amortisation Expense\n       Cr. Film & Content Asset\n   Amortised over time proportional to individual-film-forecast revenue.",
  },
  6: {
    leftBoxTitle: "THEATRICAL VALUE CHAIN DYNAMICS",
    leftBoxText:
      "Value Chain Flow:\nPRODUCER (Creates IP) → DISTRIBUTOR (Licensing) → EXHIBITOR (PVR INOX) → PATRON\n\nGross Box Office is NOT Exhibitor Revenue:\nTicket collections include GST (12% or 18%). Net box office is divided between distributor and cinema exhibitor via sliding scale theatrical agreements.",
    rightBoxTitle: "PVR INOX CORE REVENUE PILLARS",
    rightBoxText:
      "1. Net Ticket Admissions (~50-55% mix): Split with distributors.\n2. Food & Beverage (~28-32% mix): Exceptional 70-75% gross margin without distributor sharing.\n3. Screen Advertising (~8-10% mix): >80% EBITDA margin brand integration.\n4. Convenience & Digital Fees (~5-7% mix): Ticketing convenience charges.",
  },
  7: {
    leftBoxTitle: "PVR INOX LIMITED CONSOLIDATED BALANCE SHEET\n(Audited as of 31 March 2025)",
    leftBoxText:
      "• Equity Share Capital: ₹ 982 Million\n• Other Equity (Reserves & Surplus): ₹ 69,265 Million\n• Total Shareholders' Equity: ₹ 70,247 Million\n• Long-Term Borrowings: ₹ 9,198 Million\n• Current Liabilities: ₹ 24,977 Million\n• Non-Current Assets (inc. Right-of-Use): ₹ 145,920 Million\n• Current Assets: ₹ 10,810 Million\n• Total Balance Sheet Size: ₹ 156,730 Million",
    rightBoxTitle: "BALANCE SHEET ACCOUNTING INSIGHTS (IND AS 116)",
    rightBoxText:
      "1. Leasehold Capitalisation (Ind AS 116):\n   Multiplex leases in prime commercial malls are capitalised as Right-of-Use assets and corresponding lease liabilities.\n\n2. Capital Intensive Scale:\n   ₹145.9B in non-current assets reflects massive projection hardware, Dolby Atmos sound systems, and physical auditoriums.\n\n3. Balance Sheet ≠ Profit:\n   Heavy depreciation on ROU assets can produce net quarterly losses even when cash EBITDA remains healthy.",
  },
  8: {
    leftBoxTitle: "DHURANDHAR: THE BOX OFFICE WATERFALL",
    leftBoxText:
      "Theatrical Waterfall Estimations (The Week Feb 2026):\n• Worldwide Gross Box Office: ~₹ 1,250 Crore\n• India Gross Box Office: ~₹ 950 Crore\n• Less: Theatre Exhibitor Share: -~₹ 300 Crore\n• India Net Distributor Collection: ~₹ 650 Crore\n• Retained Producer Theatrical Recovery: Subject to distributor recoupment tiers",
    rightBoxTitle: "PRODUCER PROFIT FORMULA & AUDIT LESSONS",
    rightBoxText:
      "PRODUCER PROFIT =\nRecognised Producer Share (Theatrical + OTT + Satellite + Music)\n− Capitalised Production Costs\n− Prints & Advertising (P&A)\n− Financing Charges & Talent Profit Participations\n\nCRITICAL AUDITING REALITY:\nGross box office is not producer profit. Private production and talent contracts are confidential; exact profit cannot be deduced from gross box office alone.",
  },
  9: {
    leftBoxTitle: "HOLLYWOOD STUDIO MODEL (WBD)",
    leftBoxText:
      "• Primary Engine: Global IP licensing, SVOD streaming platforms, interactive games, and home entertainment.\n• Content Library: US$ 14.6B+ capitalised content assets amortised across multi-decade lifecycles.\n• Merchandise: Third-party licensee deals (MINISO, Ferrero) producing high-margin perpetual royalty cash flows.\n• Goal: Maximise franchise lifetime equity across global media.",
    rightBoxTitle: "INDIAN EXHIBITOR MODEL (PVR INOX)",
    rightBoxText:
      "• Primary Engine: Theatrical ticket admissions, high-margin F&B (popcorn & cola), and screen advertising.\n• Asset Base: ₹156.7B in physical multiplexes, laser projectors, screens, and mall leasehold commitments.\n• Merchandise: Perishable food & beverage inventory; limited consumer merchandise exposure.\n• Goal: Maximise Spend Per Head (SPH) and auditorium occupancy.",
  },
  10: {
    leftBoxTitle: "THE 3 PRINCIPLES OF ENTERTAINMENT ACCOUNTING",
    leftBoxText:
      "1. Analyse Revenue by Specific Stream:\n   Never conflate gross theatrical box office with studio revenue or exhibitor income.\n\n2. Track Capitalisation & Matching:\n   Distinguish capitalised content assets and inventory from direct operating expenses.\n\n3. Evaluate Complete Solvency:\n   True corporate value is measured through balance sheets and cash flows, not viral box office headlines.",
    rightBoxTitle: "REGULATORY & AUDIT VERIFICATION ARCHIVE",
    rightBoxText:
      "Primary Regulatory Sources:\n• SEC EDGAR: Scholastic Corp Form 10-K (FY2008 & FY2017)\n• SEC EDGAR: Warner Bros. Discovery Form 10-K (FY2025)\n• NSE India: PVR INOX Ltd FY2024-25 Annual Report\n• MCA India: Ind AS 115 & Ind AS 116 Pronouncements\n• WIPO: Harry Potter and the IP Bonanza\n• The Week: Dhurandhar Box Office Audit (Feb 2026)",
  },
};

export interface CreateSlidesResult {
  presentationId: string;
  presentationUrl: string;
  embedUrl: string;
}

/**
 * Creates a complete 10-slide presentation in Google Slides using the Google Slides API.
 */
export async function createGoogleSlidesPresentation(
  accessToken: string,
  onProgress?: (message: string, percent: number) => void
): Promise<CreateSlidesResult> {
  onProgress?.("Initialising presentation in Google Slides...", 10);

  // 1. Create a blank presentation
  const createRes = await fetch("https://slides.googleapis.com/v1/presentations", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      title: "Beyond the Box Office: How Hollywood and Indian Cinema Generate Revenue Beyond Theatres",
    }),
  });

  if (!createRes.ok) {
    const errorData = await createRes.json().catch(() => ({}));
    console.error("Failed to create Google Slides presentation:", errorData);
    throw new Error(
      errorData?.error?.message ||
        `Google Slides API error: Failed to create presentation (${createRes.status})`
    );
  }

  const presentation = await createRes.json();
  const presentationId = presentation.presentationId;
  const initialSlideId = presentation.slides?.[0]?.objectId;

  onProgress?.("Generating 10 custom 16:9 widescreen slides...", 25);

  const requests: any[] = [];
  const timestamp = Date.now().toString(36);

  // Generate batch requests for all 10 slides
  for (let i = 0; i < SLIDES_META.length; i++) {
    const meta = SLIDES_META[i];
    const content = SLIDES_CONTENT[meta.id];
    const pfx = `s${i + 1}_${timestamp}_`;
    const slideObjectId = `${pfx}slide`;

    // 1. Create Blank Slide
    requests.push({
      createSlide: {
        objectId: slideObjectId,
        slideLayoutReference: {
          predefinedLayout: "BLANK",
        },
      },
    });

    // 2. Set Slide Background to Obsidian Dark (#101012) using UpdatePagePropertiesRequest
    requests.push({
      updatePageProperties: {
        objectId: slideObjectId,
        pageProperties: {
          pageBackgroundFill: {
            solidFill: {
              color: {
                rgbColor: COLOR_OBSIDIAN,
              },
            },
          },
        },
        fields: "pageBackgroundFill.solidFill.color",
      },
    });

    // 3. Category & Number Header Text Box
    const headerShapeId = `${pfx}hdr`;
    requests.push({
      createShape: {
        objectId: headerShapeId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: slideObjectId,
          size: {
            width: { magnitude: 640, unit: "PT" },
            height: { magnitude: 24, unit: "PT" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 40,
            translateY: 20,
            unit: "PT",
          },
        },
      },
    });
    requests.push({
      insertText: {
        objectId: headerShapeId,
        insertionIndex: 0,
        text: `SLIDE ${meta.slideNumber}  |  ${meta.category}`,
      },
    });
    requests.push({
      updateTextStyle: {
        objectId: headerShapeId,
        textRange: { type: "ALL" },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: COLOR_GOLD } },
          fontSize: { magnitude: 9.5, unit: "PT" },
          bold: true,
          fontFamily: "Arial",
        },
        fields: "foregroundColor,fontSize,bold,fontFamily",
      },
    });

    // 4. Main Slide Title
    const titleShapeId = `${pfx}title`;
    requests.push({
      createShape: {
        objectId: titleShapeId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: slideObjectId,
          size: {
            width: { magnitude: 640, unit: "PT" },
            height: { magnitude: 36, unit: "PT" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 40,
            translateY: 42,
            unit: "PT",
          },
        },
      },
    });
    requests.push({
      insertText: {
        objectId: titleShapeId,
        insertionIndex: 0,
        text: meta.title,
      },
    });
    requests.push({
      updateTextStyle: {
        objectId: titleShapeId,
        textRange: { type: "ALL" },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: COLOR_PARCHMENT } },
          fontSize: { magnitude: 18, unit: "PT" },
          bold: true,
          fontFamily: "Georgia",
        },
        fields: "foregroundColor,fontSize,bold,fontFamily",
      },
    });

    // 5. Headline / Subtitle
    const subShapeId = `${pfx}sub`;
    requests.push({
      createShape: {
        objectId: subShapeId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: slideObjectId,
          size: {
            width: { magnitude: 640, unit: "PT" },
            height: { magnitude: 22, unit: "PT" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 40,
            translateY: 80,
            unit: "PT",
          },
        },
      },
    });
    requests.push({
      insertText: {
        objectId: subShapeId,
        insertionIndex: 0,
        text: `"${meta.headline}"`,
      },
    });
    requests.push({
      updateTextStyle: {
        objectId: subShapeId,
        textRange: { type: "ALL" },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: COLOR_MUTED } },
          fontSize: { magnitude: 10, unit: "PT" },
          italic: true,
          fontFamily: "Georgia",
        },
        fields: "foregroundColor,fontSize,italic,fontFamily",
      },
    });

    // 6. Left Container Box (ShapeType RECTANGLE)
    const leftBoxId = `${pfx}lbox`;
    requests.push({
      createShape: {
        objectId: leftBoxId,
        shapeType: "RECTANGLE",
        elementProperties: {
          pageObjectId: slideObjectId,
          size: {
            width: { magnitude: 305, unit: "PT" },
            height: { magnitude: 250, unit: "PT" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 40,
            translateY: 108,
            unit: "PT",
          },
        },
      },
    });
    requests.push({
      updateShapeProperties: {
        objectId: leftBoxId,
        shapeProperties: {
          shapeBackgroundFill: {
            solidFill: {
              color: { rgbColor: i % 2 === 0 ? COLOR_BURGUNDY : COLOR_MIDNIGHT },
            },
          },
          outline: {
            outlineFill: {
              solidFill: { color: { rgbColor: COLOR_GOLD } },
            },
            weight: { magnitude: 1, unit: "PT" },
          },
        },
        fields: "shapeBackgroundFill.solidFill.color,outline",
      },
    });
    const leftText = `${content.leftBoxTitle}\n\n${content.leftBoxText}`;
    requests.push({
      insertText: {
        objectId: leftBoxId,
        insertionIndex: 0,
        text: leftText,
      },
    });
    requests.push({
      updateTextStyle: {
        objectId: leftBoxId,
        textRange: { type: "ALL" },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: COLOR_PARCHMENT } },
          fontSize: { magnitude: 9.5, unit: "PT" },
          fontFamily: "Arial",
        },
        fields: "foregroundColor,fontSize,fontFamily",
      },
    });

    // 7. Right Container Box (ShapeType RECTANGLE)
    const rightBoxId = `${pfx}rbox`;
    requests.push({
      createShape: {
        objectId: rightBoxId,
        shapeType: "RECTANGLE",
        elementProperties: {
          pageObjectId: slideObjectId,
          size: {
            width: { magnitude: 320, unit: "PT" },
            height: { magnitude: 250, unit: "PT" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 360,
            translateY: 108,
            unit: "PT",
          },
        },
      },
    });
    requests.push({
      updateShapeProperties: {
        objectId: rightBoxId,
        shapeProperties: {
          shapeBackgroundFill: {
            solidFill: {
              color: { rgbColor: i % 2 === 1 ? COLOR_BURGUNDY : COLOR_MIDNIGHT },
            },
          },
          outline: {
            outlineFill: {
              solidFill: { color: { rgbColor: COLOR_GOLD } },
            },
            weight: { magnitude: 1, unit: "PT" },
          },
        },
        fields: "shapeBackgroundFill.solidFill.color,outline",
      },
    });
    const rightText = `${content.rightBoxTitle}\n\n${content.rightBoxText}`;
    requests.push({
      insertText: {
        objectId: rightBoxId,
        insertionIndex: 0,
        text: rightText,
      },
    });
    requests.push({
      updateTextStyle: {
        objectId: rightBoxId,
        textRange: { type: "ALL" },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: COLOR_PARCHMENT } },
          fontSize: { magnitude: 9.5, unit: "PT" },
          fontFamily: "Arial",
        },
        fields: "foregroundColor,fontSize,fontFamily",
      },
    });

    // 8. Footer Source Note
    const footerShapeId = `${pfx}ftr`;
    requests.push({
      createShape: {
        objectId: footerShapeId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: slideObjectId,
          size: {
            width: { magnitude: 640, unit: "PT" },
            height: { magnitude: 20, unit: "PT" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 40,
            translateY: 370,
            unit: "PT",
          },
        },
      },
    });
    requests.push({
      insertText: {
        objectId: footerShapeId,
        insertionIndex: 0,
        text: `Source Audit: ${meta.sourceNote} | Accountancy for Engineers Case Study`,
      },
    });
    requests.push({
      updateTextStyle: {
        objectId: footerShapeId,
        textRange: { type: "ALL" },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: COLOR_MUTED } },
          fontSize: { magnitude: 8, unit: "PT" },
          fontFamily: "Arial",
        },
        fields: "foregroundColor,fontSize,fontFamily",
      },
    });
  }

  // 9. If there was an initial blank default slide created by Google Slides, delete it
  if (initialSlideId) {
    requests.push({
      deleteObject: {
        objectId: initialSlideId,
      },
    });
  }

  onProgress?.("Applying layout styling and ledger financial data to Google Slides...", 65);

  // Send batch update to Google Slides API
  const batchRes = await fetch(
    `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ requests }),
    }
  );

  if (!batchRes.ok) {
    const errorData = await batchRes.json().catch(() => ({}));
    console.error("Google Slides batchUpdate failed:", errorData);
    throw new Error(
      errorData?.error?.message ||
        `Google Slides API error: Failed to populate slides (${batchRes.status})`
    );
  }

  onProgress?.("Completed! Presentation is ready in Google Slides.", 100);

  return {
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
    embedUrl: `https://docs.google.com/presentation/d/${presentationId}/embed?start=false&loop=false&delayms=3000`,
  };
}
