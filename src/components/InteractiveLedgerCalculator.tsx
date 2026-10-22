import React, { useState } from 'react';
import { Calculator, X, Sparkles, TrendingUp, DollarSign, ArrowRight } from 'lucide-react';

interface Props {
  isOpen: boolean;
  onClose: () => void;
}

export const InteractiveLedgerCalculator: React.FC<Props> = ({ isOpen, onClose }) => {
  const [activeTab, setActiveTab] = useState<'amortisation' | 'waterfall'>('amortisation');

  // Amortisation Model (ASC 926)
  const [prodBudget, setProdBudget] = useState(150); // $M
  const [ultimateRev, setUltimateRev] = useState(450); // $M
  const [year1Rev, setYear1Rev] = useState(180); // $M

  // Calculated ASC 926 Amortisation
  const amortRatio = Math.min(1, Math.max(0, year1Rev / (ultimateRev || 1)));
  const periodAmortExpense = Math.round(prodBudget * amortRatio * 10) / 10;
  const remainingAsset = Math.max(0, Math.round((prodBudget - periodAmortExpense) * 10) / 10);

  // Indian Cinema Waterfall Model
  const [grossBoxOffice, setGrossBoxOffice] = useState(800); // ₹ Cr
  const [gstRate, setGstRate] = useState(18); // %
  const [exhibitorSharePct, setExhibitorSharePct] = useState(50); // %
  const [prodAndPAMost, setProdAndPAMost] = useState(350); // ₹ Cr
  const [nonTheatricalRights, setNonTheatricalRights] = useState(250); // OTT, Satellite, Music (₹ Cr)

  // Calculations
  const netBoxOffice = Math.round((grossBoxOffice / (1 + gstRate / 100)) * 10) / 10;
  const exhibitorCut = Math.round(netBoxOffice * (exhibitorSharePct / 100) * 10) / 10;
  const distributorGrossShare = Math.round((netBoxOffice - exhibitorCut) * 10) / 10;
  const totalProducerIncome = Math.round((distributorGrossShare * 0.9 + nonTheatricalRights) * 10) / 10; // after ~10% distribution fee
  const netProducerProfit = Math.round((totalProducerIncome - prodAndPAMost) * 10) / 10;

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-[#101012]/90 backdrop-blur-md flex items-center justify-center p-4">
      <div className="w-full max-w-4xl bg-[#19171B] border border-[#C6A56B]/40 rounded-xl shadow-2xl p-6 overflow-hidden flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="flex items-center justify-between pb-4 border-b border-[#C6A56B]/20">
          <div className="flex items-center space-x-2">
            <Calculator className="w-5 h-5 text-[#C6A56B]" />
            <h3 className="font-serif text-lg text-[#F0E5D0] font-semibold">
              The Ledger Simulator · Entertainment Financial Modeling
            </h3>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded text-[#AAA5A0] hover:text-[#F0E5D0] hover:bg-[#521F2B]/50 transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab switch */}
        <div className="flex space-x-2 mt-4 pb-2 border-b border-[#C6A56B]/10">
          <button
            onClick={() => setActiveTab('amortisation')}
            className={`px-4 py-2 rounded-t text-xs font-serif font-semibold transition ${
              activeTab === 'amortisation'
                ? 'bg-[#521F2B] text-[#F0E5D0] border-b-2 border-[#C6A56B]'
                : 'text-[#AAA5A0] hover:text-[#F0E5D0]'
            }`}
          >
            1. Hollywood Content Amortisation (ASC 926)
          </button>
          <button
            onClick={() => setActiveTab('waterfall')}
            className={`px-4 py-2 rounded-t text-xs font-serif font-semibold transition ${
              activeTab === 'waterfall'
                ? 'bg-[#521F2B] text-[#F0E5D0] border-b-2 border-[#C6A56B]'
                : 'text-[#AAA5A0] hover:text-[#F0E5D0]'
            }`}
          >
            2. Indian Box Office Waterfall & Producer Recovery
          </button>
        </div>

        {/* Content body */}
        <div className="flex-1 overflow-y-auto py-4">
          {activeTab === 'amortisation' ? (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Inputs */}
              <div className="space-y-4 bg-[#101012]/60 p-4 rounded-lg border border-[#C6A56B]/20 text-xs">
                <span className="font-mono text-[#C6A56B] uppercase font-bold block">
                  Interactive Assumptions
                </span>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-[#AAA5A0]">Film Production Cost (Capitalised Asset):</span>
                    <span className="font-mono text-[#F0E5D0] font-bold">${prodBudget}M</span>
                  </div>
                  <input
                    type="range"
                    min="20"
                    max="400"
                    step="5"
                    value={prodBudget}
                    onChange={(e) => setProdBudget(Number(e.target.value))}
                    className="w-full accent-[#C6A56B]"
                  />
                </div>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-[#AAA5A0]">Projected Ultimate Revenues (Lifetime):</span>
                    <span className="font-mono text-[#F0E5D0] font-bold">${ultimateRev}M</span>
                  </div>
                  <input
                    type="range"
                    min="100"
                    max="1200"
                    step="25"
                    value={ultimateRev}
                    onChange={(e) => setUltimateRev(Number(e.target.value))}
                    className="w-full accent-[#C6A56B]"
                  />
                </div>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-[#AAA5A0]">Year 1 Recognized Revenue:</span>
                    <span className="font-mono text-[#F0E5D0] font-bold">${year1Rev}M</span>
                  </div>
                  <input
                    type="range"
                    min="10"
                    max={ultimateRev}
                    step="10"
                    value={Math.min(year1Rev, ultimateRev)}
                    onChange={(e) => setYear1Rev(Number(e.target.value))}
                    className="w-full accent-[#C6A56B]"
                  />
                </div>

                <div className="p-3 bg-[#19171B] rounded border border-[#C6A56B]/20 text-[11px] text-[#AAA5A0] leading-relaxed">
                  <strong className="text-[#F0E5D0]">ASC 926 Rule:</strong> Production spending is not expensed when cash is paid. It is capitalised to the balance sheet as an asset, then expensed in proportion to actual revenue vs forecasted ultimate revenues.
                </div>
              </div>

              {/* Calculations & Journal Entry */}
              <div className="space-y-4 flex flex-col justify-between">
                <div className="bg-[#521F2B]/30 border border-[#C6A56B]/30 rounded-lg p-4 space-y-3">
                  <span className="font-mono text-[#C6A56B] uppercase font-bold text-xs block">
                    Calculated Accounting Outputs
                  </span>

                  <div className="grid grid-cols-2 gap-3 text-xs">
                    <div className="bg-[#101012] p-2.5 rounded border border-[#C6A56B]/20">
                      <span className="text-[10px] text-[#AAA5A0] block">Amortisation Ratio:</span>
                      <span className="font-mono text-base font-bold text-[#C6A56B]">
                        {(amortRatio * 100).toFixed(1)}%
                      </span>
                    </div>

                    <div className="bg-[#101012] p-2.5 rounded border border-[#C6A56B]/20">
                      <span className="text-[10px] text-[#AAA5A0] block">Period Amortisation:</span>
                      <span className="font-mono text-base font-bold text-red-300">
                        ${periodAmortExpense}M
                      </span>
                    </div>

                    <div className="col-span-2 bg-[#101012] p-2.5 rounded border border-[#829C80]/30">
                      <span className="text-[10px] text-[#AAA5A0] block">Remaining Film Asset on Balance Sheet:</span>
                      <span className="font-mono text-lg font-bold text-[#829C80]">
                        ${remainingAsset} Million
                      </span>
                    </div>
                  </div>
                </div>

                {/* Journal Entry */}
                <div className="bg-[#101012] border border-[#C6A56B]/30 rounded-lg p-3 font-mono text-xs text-[#F0E5D0]">
                  <span className="text-[9px] text-[#C6A56B] uppercase block mb-1">Generated Journal Entry</span>
                  <div className="flex justify-between text-[#829C80]">
                    <span>Dr. Content Amortisation Expense (P&amp;L)</span>
                    <span>${periodAmortExpense}M</span>
                  </div>
                  <div className="flex justify-between pl-4 text-[#AAA5A0]">
                    <span>Cr. Film Production Asset (B/S)</span>
                    <span>${periodAmortExpense}M</span>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Inputs */}
              <div className="space-y-4 bg-[#101012]/60 p-4 rounded-lg border border-[#C6A56B]/20 text-xs">
                <span className="font-mono text-[#C6A56B] uppercase font-bold block">
                  Theatrical Box Office Waterfall
                </span>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-[#AAA5A0]">India Theatrical Gross:</span>
                    <span className="font-mono text-[#F0E5D0] font-bold">₹ {grossBoxOffice} Cr</span>
                  </div>
                  <input
                    type="range"
                    min="50"
                    max="1500"
                    step="25"
                    value={grossBoxOffice}
                    onChange={(e) => setGrossBoxOffice(Number(e.target.value))}
                    className="w-full accent-[#C6A56B]"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <span className="text-[#AAA5A0] block mb-1">GST Rate:</span>
                    <select
                      value={gstRate}
                      onChange={(e) => setGstRate(Number(e.target.value))}
                      className="w-full bg-[#19171B] border border-[#C6A56B]/30 rounded p-1.5 text-xs text-[#F0E5D0]"
                    >
                      <option value={12}>12% (Ticket &le; ₹100)</option>
                      <option value={18}>18% (Ticket &gt; ₹100)</option>
                    </select>
                  </div>
                  <div>
                    <span className="text-[#AAA5A0] block mb-1">Exhibitor Share:</span>
                    <select
                      value={exhibitorSharePct}
                      onChange={(e) => setExhibitorSharePct(Number(e.target.value))}
                      className="w-full bg-[#19171B] border border-[#C6A56B]/30 rounded p-1.5 text-xs text-[#F0E5D0]"
                    >
                      <option value={45}>45% (Blockbuster Wk 1)</option>
                      <option value={50}>50% (Standard 50:50)</option>
                      <option value={55}>55% (Extended run)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-[#AAA5A0]">Combined Production + P&amp;A:</span>
                    <span className="font-mono text-[#F0E5D0] font-bold">₹ {prodAndPAMost} Cr</span>
                  </div>
                  <input
                    type="range"
                    min="50"
                    max="600"
                    step="25"
                    value={prodAndPAMost}
                    onChange={(e) => setProdAndPAMost(Number(e.target.value))}
                    className="w-full accent-[#C6A56B]"
                  />
                </div>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-[#AAA5A0]">Non-Theatrical Ancillary (OTT + Sat):</span>
                    <span className="font-mono text-[#829C80] font-bold">₹ {nonTheatricalRights} Cr</span>
                  </div>
                  <input
                    type="range"
                    min="0"
                    max="500"
                    step="25"
                    value={nonTheatricalRights}
                    onChange={(e) => setNonTheatricalRights(Number(e.target.value))}
                    className="w-full accent-[#C6A56B]"
                  />
                </div>
              </div>

              {/* Waterfall Steps */}
              <div className="bg-[#521F2B]/30 border border-[#C6A56B]/30 rounded-lg p-4 space-y-2.5 text-xs">
                <span className="font-mono text-[#C6A56B] uppercase font-bold block">
                  Waterfall Accounting Breakdown
                </span>

                <div className="space-y-1.5">
                  <div className="flex justify-between bg-[#101012] p-2 rounded">
                    <span className="text-[#AAA5A0]">1. India Gross Collection:</span>
                    <span className="font-mono text-[#F0E5D0] font-bold">₹ {grossBoxOffice} Cr</span>
                  </div>
                  <div className="flex justify-between bg-[#101012] p-2 rounded">
                    <span className="text-[#AAA5A0]">2. Net Collections (Post GST):</span>
                    <span className="font-mono text-[#F0E5D0]">₹ {netBoxOffice} Cr</span>
                  </div>
                  <div className="flex justify-between bg-[#101012] p-2 rounded text-red-300">
                    <span>3. Retained by Exhibitors (PVR INOX):</span>
                    <span className="font-mono">- ₹ {exhibitorCut} Cr</span>
                  </div>
                  <div className="flex justify-between bg-[#101012] p-2 rounded text-[#C6A56B]">
                    <span>4. Distributor Net Theatrical Share:</span>
                    <span className="font-mono font-bold">₹ {distributorGrossShare} Cr</span>
                  </div>
                  <div className="flex justify-between bg-[#101012] p-2 rounded text-[#829C80]">
                    <span>5. Ancillary Recovery (OTT/Sat/Music):</span>
                    <span className="font-mono">+ ₹ {nonTheatricalRights} Cr</span>
                  </div>
                </div>

                <div
                  className={`p-3 rounded border text-center font-mono mt-3 ${
                    netProducerProfit >= 0
                      ? 'bg-[#101012] border-[#829C80] text-[#829C80]'
                      : 'bg-[#521F2B] border-red-500 text-red-200'
                  }`}
                >
                  <span className="text-[10px] block text-[#AAA5A0] uppercase">
                    Estimated Net Producer Profit / (Loss)
                  </span>
                  <span className="text-xl font-bold">
                    {netProducerProfit >= 0 ? `+ ₹ ${netProducerProfit} Cr` : `- ₹ ${Math.abs(netProducerProfit)} Cr`}
                  </span>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="pt-3 border-t border-[#C6A56B]/20 flex justify-between items-center text-[11px] text-[#AAA5A0]">
          <span>Designed for classroom exploration · US GAAP ASC 926 &amp; Indian Theatrical Waterfalls</span>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-[#C6A56B] text-[#101012] font-semibold rounded text-xs hover:bg-[#e0c38c] transition"
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
};
