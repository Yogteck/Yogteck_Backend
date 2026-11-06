/**
 * Convert number into Indian Currency Words (English)
 * Example: 4490 -> FOUR THOUSAND FOUR HUNDRED AND NINETY RUPEES ONLY
 */
function numberToWords(amount) {
  const num = Math.round(Number(amount || 0) * 100) / 100;
  if (isNaN(num) || num === 0) return 'ZERO RUPEES ONLY';

  const [rupeesPart, paisePart] = num.toFixed(2).split('.').map(n => parseInt(n, 10));

  const ones = [
    '', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE',
    'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN',
    'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'
  ];

  const tens = [
    '', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'
  ];

  function convertTwoDigits(n) {
    if (n < 20) return ones[n];
    const unit = n % 10;
    return (tens[Math.floor(n / 10)] + (unit ? ' ' + ones[unit] : '')).trim();
  }

  function convertThreeDigits(n) {
    let str = '';
    const hundred = Math.floor(n / 100);
    const rem = n % 100;
    if (hundred > 0) {
      str += ones[hundred] + ' HUNDRED';
    }
    if (rem > 0) {
      if (str !== '') str += ' AND ';
      str += convertTwoDigits(rem);
    }
    return str.trim();
  }

  function convertToIndianWords(n) {
    if (n === 0) return '';

    let result = '';

    // Crores (>= 1,00,00,000)
    const crore = Math.floor(n / 10000000);
    let rem = n % 10000000;
    if (crore > 0) {
      result += convertToIndianWords(crore) + ' CRORE ';
    }

    // Lakhs (>= 1,00,000)
    const lakh = Math.floor(rem / 100000);
    rem = rem % 100000;
    if (lakh > 0) {
      result += convertTwoDigits(lakh) + ' LAKH ';
    }

    // Thousands (>= 1,000)
    const thousand = Math.floor(rem / 1000);
    rem = rem % 1000;
    if (thousand > 0) {
      result += convertTwoDigits(thousand) + ' THOUSAND ';
    }

    // Hundreds and below (< 1,000)
    if (rem > 0) {
      result += convertThreeDigits(rem);
    }

    return result.trim();
  }

  let words = convertToIndianWords(rupeesPart) + ' RUPEES';

  if (paisePart > 0) {
    words += ' AND ' + convertTwoDigits(paisePart) + ' PAISE';
  }

  words += ' ONLY';
  return words.replace(/\s+/g, ' ').trim();
}

module.exports = { numberToWords };

