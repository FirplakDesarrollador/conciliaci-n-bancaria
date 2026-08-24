import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });
import { sapClient } from './src/lib/sap/service-layer';
import * as fs from 'fs';

async function queryDoc(docNum: number) {
    console.log(`Consultando SAP para ${docNum}...`);
    try {
        const resIn = await sapClient.request(`/IncomingPayments?$filter=DocNum eq ${docNum}`);
        const dataIn = await resIn.json();
        if (dataIn.value && dataIn.value.length > 0) {
            return { type: 'IncomingPayment', data: dataIn.value[0] };
        } else {
            const resOut = await sapClient.request(`/VendorPayments?$filter=DocNum eq ${docNum}`);
            const dataOut = await resOut.json();
            if (dataOut.value && dataOut.value.length > 0) {
                return { type: 'VendorPayment', data: dataOut.value[0] };
            }
        }
    } catch (e) {
        console.error(`Error querying ${docNum}:`, e);
    }
    return { type: 'NotFound', data: null };
}

async function main() {
    const docs = [80832, 80834, 80835, 80847, 80849, 80828, 80837, 20043385, 20043386, 20043387, 20043388];
    const results: any = {};
    for (const doc of docs) {
        results[doc] = await queryDoc(doc);
    }
    fs.writeFileSync('sap_analysis_aug20.json', JSON.stringify(results, null, 2));
    console.log("Análisis completado y guardado en sap_analysis_aug20.json");
}

main().catch(console.error);
