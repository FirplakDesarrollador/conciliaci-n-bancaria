import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });
import { sapClient } from './src/lib/sap/service-layer';

async function main() {
    const resOut = await sapClient.request(`/VendorPayments?$filter=DocNum eq 20043323`);
    const dataOut = await resOut.json();
    console.log(JSON.stringify(dataOut.value?.[0], null, 2));
}

main().catch(console.error);
