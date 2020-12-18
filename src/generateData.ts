const paymentTypes = ["선불", "후불현금(VAN)", "후불카드(VAN"];

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

const generateData = (length: number = 10000): Data[] =>
  Array.apply(undefined, Array(length)).map(() => ({
    orderNumber: `20200101010101001#${Math.floor(Math.random() * 10000)}`,
    date: "2020-01-01",
    clientCharge:
      Math.random() > 0.8 ? 15000 + Math.floor(Math.random() * 10) * 1000 : 0,
    cancelFee: 0,
    deliveryFee: 3500 + Math.floor(Math.random() * 100) * 10,
    creditCardCharge:
      Math.random() > 0.5 ? 15000 + Math.floor(Math.random() * 10) * 1000 : 0,
    creditCardFee: 0,
    creditCardFeeRate: 0,
    cashReceiptId: undefined,
    paymentType: paymentTypes[Math.floor((Math.random() * 3) % 3)],
  }));

export default generateData;
