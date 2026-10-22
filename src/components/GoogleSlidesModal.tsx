import React, { useState } from 'react';
import {
  X,
  ExternalLink,
  CheckCircle2,
  AlertCircle,
  Loader2,
  FileSpreadsheet,
  Tv,
  Layers,
  Sparkles,
} from 'lucide-react';
import { createGoogleSlidesPresentation, CreateSlidesResult } from '../services/googleSlides';
import { SLIDES_META } from '../data/slidesData';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  accessToken: string | null;
  onRequireAuth: () => void;
}

export const GoogleSlidesModal: React.FC<Props> = ({
  isOpen,
  onClose,
  accessToken,
  onRequireAuth,
}) => {
  const [confirmed, setConfirmed] = useState(false);
  const [isCreating, setIsCreating] = useState(false);
  const [progressMsg, setProgressMsg] = useState('');
  const [progressPct, setProgressPct] = useState(0);
  const [result, setResult] = useState<CreateSlidesResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [showEmbedPreview, setShowEmbedPreview] = useState(false);

  if (!isOpen) return null;

  const handleCreate = async () => {
    if (!accessToken) {
      onRequireAuth();
      return;
    }

    setIsCreating(true);
    setError(null);
    setProgressPct(10);
    setProgressMsg('Connecting to Google Slides API...');

    try {
      const res = await createGoogleSlidesPresentation(
        accessToken,
        (msg, pct) => {
          setProgressMsg(msg);
          setProgressPct(pct);
        }
      );
      setResult(res);
      setIsCreating(false);
    } catch (err: any) {
      console.error(err);
      setError(
        err.message || 'Failed to generate presentation in Google Slides. Please verify account permissions.'
      );
      setIsCreating(false);
    }
  };

  const handleReset = () => {
    setResult(null);
    setError(null);
    setConfirmed(false);
    setIsCreating(false);
    setProgressPct(0);
    setProgressMsg('');
    setShowEmbedPreview(false);
  };

  return (
    <div className="fixed inset-0 z-50 bg-[#101012]/90 backdrop-blur-md flex items-center justify-center p-4">
      <div className="w-full max-w-2xl bg-[#19171B] border border-[#C6A56B]/40 rounded-xl shadow-2xl p-6 overflow-hidden flex flex-col">
        {/* Header */}
        <div className="flex items-center justify-between pb-4 border-b border-[#C6A56B]/20">
          <div className="flex items-center space-x-2">
            <svg className="w-6 h-6" viewBox="0 0 24 24" fill="none">
              <path
                d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10H7v-2h10v2zm0-4H7V7h10v2z"
                fill="#F4B400"
              />
            </svg>
            <h3 className="font-serif text-lg text-[#F0E5D0] font-semibold">
              Export to Google Slides
            </h3>
          </div>
          <button
            onClick={() => {
              handleReset();
              onClose();
            }}
            className="p-1 rounded text-[#AAA5A0] hover:text-[#F0E5D0] hover:bg-[#521F2B]/50 transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="py-5">
          {!result ? (
            <div>
              {/* User Confirmation Card */}
              <div className="bg-[#101012] border border-[#C6A56B]/30 rounded-lg p-5 space-y-4">
                <div className="flex items-start space-x-3">
                  <div className="p-2 rounded bg-[#C6A56B]/10 border border-[#C6A56B]/30 text-[#C6A56B]">
                    <Layers className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="font-serif text-sm text-[#F0E5D0] font-bold">
                      Create 10-Slide Deck in Your Google Slides
                    </h4>
                    <p className="text-xs text-[#AAA5A0] mt-1 leading-relaxed">
                      This action will generate a new 16:9 widescreen presentation titled{" "}
                      <strong className="text-[#F0E5D0]">
                        "Beyond the Box Office: How Hollywood and Indian Cinema Generate Revenue Beyond Theatres"
                      </strong>{" "}
                      in your connected Google account.
                    </p>
                  </div>
                </div>

                <div className="border-t border-[#C6A56B]/15 pt-3 space-y-2 text-xs">
                  <span className="font-semibold text-[#C6A56B] uppercase font-mono text-[10px] block">
                    Deck Content Specification:
                  </span>
                  <div className="grid grid-cols-2 gap-2 text-[11px] text-[#AAA5A0]">
                    <div className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-[#829C80]" />
                      <span>10 Widescreen 16:9 Custom Slides</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-[#829C80]" />
                      <span>The Cinematic Ledger Obsidian Theme</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-[#829C80]" />
                      <span>WBD &amp; PVR INOX Audited Tables</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-[#829C80]" />
                      <span>ASC 606 &amp; Ind AS 116 Journal Entries</span>
                    </div>
                  </div>
                </div>

                {!accessToken && (
                  <div className="p-3 bg-[#521F2B]/30 border border-[#C6A56B]/40 rounded text-xs flex items-center justify-between">
                    <span className="text-[#F0E5D0]">Google account authentication required to proceed.</span>
                    <button
                      onClick={onRequireAuth}
                      className="px-3 py-1 bg-[#C6A56B] text-[#101012] font-semibold rounded text-xs hover:bg-[#e0c38c] transition"
                    >
                      Sign In Now
                    </button>
                  </div>
                )}

                {/* Progress Indicator */}
                {isCreating && (
                  <div className="space-y-2 pt-2">
                    <div className="flex justify-between text-xs text-[#C6A56B]">
                      <span className="flex items-center gap-2">
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        {progressMsg}
                      </span>
                      <span className="font-mono">{progressPct}%</span>
                    </div>
                    <div className="w-full bg-[#19171B] h-2 rounded-full overflow-hidden border border-[#C6A56B]/30">
                      <div
                        className="bg-gradient-to-r from-[#521F2B] via-[#C6A56B] to-[#829C80] h-full transition-all duration-300"
                        style={{ width: `${progressPct}%` }}
                      />
                    </div>
                  </div>
                )}

                {/* Error Banner */}
                {error && (
                  <div className="p-3 bg-[#521F2B] border border-red-500 rounded text-xs text-red-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <AlertCircle className="w-4 h-4 shrink-0 text-red-400" />
                      <span>{error}</span>
                    </div>
                    {(error.toLowerCase().includes('auth') || error.toLowerCase().includes('token') || error.toLowerCase().includes('401') || error.toLowerCase().includes('permission')) && (
                      <button
                        onClick={onRequireAuth}
                        className="px-2.5 py-1 bg-[#C6A56B] text-[#101012] font-bold rounded text-[11px] shrink-0 hover:bg-[#e0c38c] transition self-start sm:self-auto"
                      >
                        Re-authenticate Google
                      </button>
                    )}
                  </div>
                )}
              </div>

              {/* Action Buttons */}
              <div className="flex justify-end space-x-3 mt-5">
                <button
                  onClick={() => {
                    handleReset();
                    onClose();
                  }}
                  disabled={isCreating}
                  className="px-4 py-2 rounded text-xs text-[#AAA5A0] hover:text-[#F0E5D0] border border-[#C6A56B]/20 hover:border-[#C6A56B]/40 transition"
                >
                  Cancel
                </button>
                <button
                  onClick={handleCreate}
                  disabled={isCreating}
                  className="flex items-center space-x-2 px-5 py-2 rounded text-xs font-semibold text-[#101012] bg-[#F4B400] hover:bg-[#fcd04d] transition shadow-lg disabled:opacity-50"
                >
                  {isCreating ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>Creating in Google Slides...</span>
                    </>
                  ) : (
                    <>
                      <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none">
                        <path
                          d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10H7v-2h10v2zm0-4H7V7h10v2z"
                          fill="#101012"
                        />
                      </svg>
                      <span>Confirm &amp; Create in Google Slides</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          ) : (
            /* Success State */
            <div className="space-y-4">
              <div className="bg-[#521F2B]/30 border border-[#829C80] rounded-lg p-5 text-center space-y-3">
                <div className="w-12 h-12 rounded-full bg-[#829C80]/20 border border-[#829C80] mx-auto flex items-center justify-center text-[#829C80]">
                  <CheckCircle2 className="w-6 h-6" />
                </div>
                <h4 className="font-serif text-lg text-[#F0E5D0] font-bold">
                  Presentation Successfully Created!
                </h4>
                <p className="text-xs text-[#AAA5A0] max-w-md mx-auto">
                  Your 10-slide case study is live in your Google Slides account. You can now edit, present, or share with students and colleagues.
                </p>

                <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
                  <a
                    href={result.presentationUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center space-x-2 px-5 py-2.5 rounded text-xs font-semibold text-[#101012] bg-[#C6A56B] hover:bg-[#e0c38c] transition shadow-md"
                  >
                    <span>Open in Google Slides</span>
                    <ExternalLink className="w-3.5 h-3.5" />
                  </a>

                  <button
                    onClick={() => setShowEmbedPreview(!showEmbedPreview)}
                    className="flex items-center space-x-1.5 px-4 py-2 rounded text-xs text-[#F0E5D0] bg-[#101012] border border-[#C6A56B]/40 hover:border-[#C6A56B] transition"
                  >
                    <Tv className="w-3.5 h-3.5 text-[#C6A56B]" />
                    <span>{showEmbedPreview ? 'Hide Live Embed' : 'Preview Embedded'}</span>
                  </button>
                </div>
              </div>

              {/* Embedded Player */}
              {showEmbedPreview && (
                <div className="border border-[#C6A56B]/30 rounded-lg overflow-hidden bg-black aspect-[16/9]">
                  <iframe
                    src={result.embedUrl}
                    title="Google Slides Presentation Preview"
                    width="100%"
                    height="100%"
                    allowFullScreen
                    className="border-0"
                  />
                </div>
              )}

              <div className="flex justify-between items-center text-xs text-[#AAA5A0] pt-2">
                <span>Presentation ID: <code className="font-mono text-[#C6A56B]">{result.presentationId}</code></span>
                <button
                  onClick={handleReset}
                  className="text-xs underline hover:text-[#F0E5D0]"
                >
                  Export Another Copy
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
