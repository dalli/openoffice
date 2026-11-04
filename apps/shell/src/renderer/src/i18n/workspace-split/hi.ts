import type { zh } from './zh'

export const hi = {
  splitView: 'दृश्य बाँटें',
  splitSide: 'अगल-बगल देखें',
  splitStack: 'ऊपर और नीचे देखें',
  splitUnsplit: 'विभाजन हटाएँ',
  splitUnavailable: 'मुखपृष्ठ, सहायता और प्रस्तुति दृश्य नहीं बाँटे जा सकते।',
  splitSingleCopy: 'एक दस्तावेज़: अलग प्रति खुलेगी; बदलाव समन्वयित नहीं होंगे।',
  splitLeftDocument: 'बायाँ दस्तावेज़',
  splitRightDocument: 'दायाँ दस्तावेज़',
  splitTopDocument: 'ऊपरी दस्तावेज़',
  splitBottomDocument: 'निचला दस्तावेज़',
} satisfies Record<keyof typeof zh, string>
