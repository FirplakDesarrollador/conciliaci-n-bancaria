import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });
import { sapClient } from './src/lib/sap/service-layer';

async function main() {
    console.log("Consultando SAP para 80639...");
    const resIn = await sapClient.request(`/IncomingPayments?$filter=DocNum eq 80639`);
    const dataIn = await resIn.json();
    if (dataIn.value && dataIn.value.length > 0) {
        console.log("Encontrado en IncomingPayments:");
        console.log(JSON.stringify(dataIn.value[0], null, 2));
    } else {
        const resOut = await sapClient.request(`/VendorPayments?$filter=DocNum eq 80639`);
        const dataOut = await resOut.json();
        if (dataOut.value && dataOut.value.length > 0) {
            console.log("Encontrado en VendorPayments:");
            console.log(JSON.stringify(dataOut.value[0], null, 2));
        } else {
            console.log("No se encontró el documento 80639 en SAP.");
        }
    }
}

main().catch(console.error);
