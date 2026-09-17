import "server-only";

// Días festivos para Colombia (YYYY-MM-DD)
const HOLIDAYS = [
  '2026-01-01', '2026-01-12', '2026-03-23', '2026-04-02', '2026-04-03',
  '2026-05-01', '2026-05-18', '2026-06-08', '2026-06-15', '2026-06-29',
  '2026-07-13', '2026-07-20', '2026-08-07', '2026-08-17', '2026-10-12',
  '2026-11-02', '2026-11-16', '2026-12-08', '2026-12-25'
];

/**
 * Fechas (YYYY-MM-DD, zona Bogota) que se consultan en SAP para el cruce
 * diario: desde 15 dias antes del inicio de la semana/puente hasta hoy.
 * Usada tanto por el dashboard (page.tsx) como por el cron de auto-sync,
 * para que ambos consulten exactamente la misma ventana.
 */
export function getTargetDates(): string[] {
  const now = new Date();
  const getBogotaDate = (daysOffset: number) => {
    const d = new Date(now.getTime() + daysOffset * 24 * 60 * 60 * 1000);
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Bogota' }).format(d);
  };

  const todayBogotaStr = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Bogota', weekday: 'short' }).format(now);
  const yesterdayBogotaStr = getBogotaDate(-1);

  const dates: string[] = [];
  let minOffset = -1;

  if (todayBogotaStr === 'Mon') {
    minOffset = -3;
  } else if (todayBogotaStr === 'Tue' && HOLIDAYS.includes(yesterdayBogotaStr)) {
    minOffset = -4;
  }

  for (let offset = minOffset - 15; offset <= 0; offset++) {
    dates.push(getBogotaDate(offset));
  }

  return dates;
}
