/**
 * 解析金額字串（移除非數字與小數點符號，例如 NT$152,000.00）
 */
export const parseAmount = (value: string | undefined): number => {
  if (!value) return NaN;
  const cleaned = value.toString().replace(/[^0-9.]/g, '');
  return parseFloat(cleaned);
};

/**
 * 將可能含逗號、文字的數字清理並轉成 number，失敗回傳 NaN
 */
export const parseNumber = (value: string | undefined): number => {
  if (!value) return NaN;
  const cleaned = value.toString().replace(/[^0-9.]/g, '');
  return parseFloat(cleaned);
};

/**
 * 取得客戶物件中的左 / 右耳 PTA 數值
 */
export const getEarPTA = (customer: { [key: string]: string }, earLabel: '左' | '右'): number => {
  const key = Object.keys(customer).find(k => k.includes(`${earLabel}耳`) && k.toUpperCase().includes('PTA'));
  if (!key) return NaN;
  return parseNumber(customer[key]);
};

/**
 * 取得聽損程度分類
 * 根據 WHO 分級 (或使用資料中的分類)
 */
export const getHearingLossDegree = (pta: number): string => {
  if (isNaN(pta)) return '未知';
  if (pta <= 25) return '正常';
  if (pta <= 40) return '輕度';
  if (pta <= 55) return '中度';
  if (pta <= 70) return '中重度';
  if (pta <= 90) return '重度';
  return '極重度';
};

/**
 * 解析年齡
 */
export const parseAge = (value: string | undefined, birthDateStr: string | undefined): number => {
  // If 'Age' column exists
  if (value) {
    const parsed = parseInt(value);
    if (!isNaN(parsed)) return parsed;
  }

  // If 'BirthDate' column exists calculate age
  if (birthDateStr) {
    const birthDate = new Date(birthDateStr);
    if (!isNaN(birthDate.getTime())) {
      const today = new Date();
      let age = today.getFullYear() - birthDate.getFullYear();
      const m = today.getMonth() - birthDate.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
        age--;
      }
      return age;
    }
  }

  return NaN;
};
