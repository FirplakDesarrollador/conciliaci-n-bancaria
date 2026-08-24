import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });
import { sapClient } from './src/lib/sap/service-layer';

async function main() {
    console.log("Consultando posibles ambiguos...");
    // 63900 en Bancolombia el 10 de Agosto
    const res1 = await sapClient.request(`/IncomingPayments?$filter=DocDate eq '2026-08-10' and TransferSum eq 63900`);
    const data1 = await res1.json();
    console.log(`Bancolombia 63900: Encontrados ${data1.value ? data1.value.length : 0} documentos`);
    if(data1.value) data1.value.forEach((v: any) => console.log(`  - DocNum: ${v.DocNum}, Tercero: ${v.CardName}`));

    // 8174925 en Banco de Bogota el 10 de Agosto (podría ser TransferDate o DocDate)
    const res2 = await sapClient.request(`/IncomingPayments?$filter=TransferSum eq 8174925`);
    const data2 = await res2.json();
    console.log(`Banco Bogota 8174925: Encontrados ${data2.value ? data2.value.length : 0} documentos`);
    if(data2.value) data2.value.forEach((v: any) => console.log(`  - DocNum: ${v.DocNum}, Date: ${v.DocDate}, Tercero: ${v.CardName}`));
}

main().catch(console.error);
