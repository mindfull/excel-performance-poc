const paymentTypes = ["선불", "후불카드(VAN)"];

export interface Data {
  orderNumber: string;
  date: string;
  clientCharge: number;
  cancelFee: number;
  deliveryFee: number;
  creditCardCharge: number;
  creditCardFee: number;
  creditCardFeeRate: number;
  cashReceiptId: string | undefined;
  paymentType: string;
}

export type SplittedData = Record<string, Data[]>;

const generateData = (length: number = 10000): Data[] => {
  const res = Array.apply(undefined, Array(length)).map(() => {
    const year = Math.random() < 0.001 ? "2019" : "2020";
    const month = (Math.floor(Math.random() * 12) + 1)
      .toString()
      .padStart(2, "0");
    const date = (Math.floor(Math.random() * 28) + 1)
      .toString()
      .padStart(2, "0");
    const isCash = Math.random() < 0.05;

    return {
      orderNumber: `${year}${month}${date}010101001#${Math.floor(
        Math.random() * 10000
      )}`,
      date: `${year}-${month}-${date}`,
      clientCharge:
        Math.random() > 0.8 ? 15000 + Math.floor(Math.random() * 10) * 1000 : 0,
      cancelFee: 0,
      deliveryFee: 3500 + Math.floor(Math.random() * 100) * 10,
      creditCardCharge:
        Math.random() > 0.5 ? 15000 + Math.floor(Math.random() * 10) * 1000 : 0,
      creditCardFee: 0,
      creditCardFeeRate: 0,
      cashReceiptId: isCash ? "0123456789" : undefined,
      paymentType: isCash
        ? "후불현금"
        : paymentTypes[Math.floor((Math.random() * 2) % 2)],
    };
  });
  return res.sort((a, b) => a.date.localeCompare(b.date));
};

export default generateData;
