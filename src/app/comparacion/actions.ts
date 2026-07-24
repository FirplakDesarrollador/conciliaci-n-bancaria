'use server';

import { redirect } from 'next/navigation';
import { createClient } from '@/lib/supabase/server';
import { runReconciliation } from '@/lib/conciliacion/engine';

export async function triggerReconciliation() {
  const supabase = await createClient();
  const { data: { user } } = await supabase.auth.getUser();

  if (!user) {
    throw new Error('Unauthorized');
  }

  try {
    const result = await runReconciliation();
    return { success: true, result };
  } catch (error: any) {
    console.error("Error en triggerReconciliation:", error);
    return { success: false, error: error.message || String(error) };
  }
}
