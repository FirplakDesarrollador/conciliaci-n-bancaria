import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });
import { sapClient } from './src/lib/sap/service-layer';

async function main() {
    console.log("Consultando SAP para 80788...");
    const resIn = await sapClient.request(`/IncomingPayments?$filter=DocNum eq 80788`);
    const dataIn = await resIn.json();
    console.log("80788:", JSON.stringify(dataIn.value?.[0], null, 2));

    const vendorDocs = [20043323, 20043325, 20043326, 20043336, 20043339, 20043348, 20043349];
    for (const doc of vendorDocs) {
        console.log(`Consultando SAP para VendorPayment ${doc}...`);
        const resOut = await sapClient.request(`/VendorPayments?$filter=DocNum eq ${doc}`);
        const dataOut = await resOut.json();
        const payment = dataOut.value?.[0];
        if (payment) {
            console.log(`VendorPayment ${doc}:`, {
                DocDate: payment.DocDate,
                TransferAccount: payment.TransferAccount,
                CashAccount: payment.CashAccount,
                CardCode: payment.CardCode,
                DocCurrency: payment.DocCurrency,
                Cancelled: payment.Cancelled,
                Remarks: payment.Remarks,
                JournalRemarks: payment.JournalRemarks
            });
        } else {
            console.log(`VendorPayment ${doc} NO ENCONTRADO`);
        }
    }
}

main().catch(console.error);
