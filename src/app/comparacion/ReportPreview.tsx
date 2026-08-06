"use client";

import React, { useState, useMemo } from 'react';
import { AgGridReact } from 'ag-grid-react';
import { ColDef, RowClassRules, ModuleRegistry, AllCommunityModule, themeQuartz } from 'ag-grid-community';

ModuleRegistry.registerModules([AllCommunityModule]);

export interface ReportItem {
  docNum: string;
  valor: number;
  banco: string;
  fecha: string;
  info: string;
  estado: "Conciliado" | "Pendiente";
}

interface ReportPreviewProps {
  pagosRecibidos: ReportItem[];
  pagosEfectuados: ReportItem[];
}

export default function ReportPreview({ pagosRecibidos, pagosEfectuados }: ReportPreviewProps) {
  const [activeTab, setActiveTab] = useState<"recibidos" | "efectuados">("recibidos");

  const rowData = activeTab === "recibidos" ? pagosRecibidos : pagosEfectuados;

  const columnDefs = useMemo<ColDef[]>(() => [
    { 
      field: "docNum", 
      headerName: "Número", 
      filter: 'agTextColumnFilter', 
      width: 150 
    },
    { 
      field: "valor", 
      headerName: "Valor", 
      filter: 'agNumberColumnFilter',
      valueFormatter: (params) => {
        return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP' }).format(params.value);
      }
    },
    { 
      field: "banco", 
      headerName: "Banco / Tercero", 
      filter: 'agTextColumnFilter',
      flex: 1
    },
    { 
      field: "fecha", 
      headerName: "Fecha", 
      filter: 'agDateColumnFilter',
      width: 150
    },
    { 
      field: "info", 
      headerName: "Descripción / Info", 
      filter: 'agTextColumnFilter',
      flex: 1.5
    },
    {
      field: "estado",
      headerName: "Estado",
      filter: 'agSetColumnFilter', // allows checking specific values
      width: 130,
      cellRenderer: (params: any) => {
        const isConciliado = params.value === "Conciliado";
        return (
          <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold ${isConciliado ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
            {params.value}
          </span>
        );
      }
    }
  ], []);

  const defaultColDef = useMemo<ColDef>(() => {
    return {
      sortable: true,
      resizable: true,
      floatingFilter: true, // Enables the Excel-like filter row under headers
    };
  }, []);

  const rowClassRules = useMemo<RowClassRules>(() => {
    return {
      // Apply yellow background slightly transparent if pending
      'bg-amber-50': (params) => params.data.estado === 'Pendiente',
    };
  }, []);

  return (
    <div className="flex flex-col h-full gap-4">
      {/* Header and Tabs */}
      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm shrink-0 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h3 className="text-lg font-bold text-slate-900">Vista Previa del Informe</h3>
          <p className="text-sm text-slate-500">
            Filtra y visualiza los datos consolidados antes de exportarlos.
          </p>
        </div>
        
        <div className="flex bg-slate-100 p-1 rounded-lg">
          <button
            onClick={() => setActiveTab("recibidos")}
            className={`px-4 py-2 rounded-md text-sm font-medium transition-all ${
              activeTab === "recibidos" 
                ? 'bg-white text-blue-700 shadow-sm' 
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Pagos Recibidos
            <span className={`ml-2 px-2 py-0.5 rounded-full text-xs ${activeTab === 'recibidos' ? 'bg-blue-100' : 'bg-slate-200'}`}>
              {pagosRecibidos.length}
            </span>
          </button>
          <button
            onClick={() => setActiveTab("efectuados")}
            className={`px-4 py-2 rounded-md text-sm font-medium transition-all ${
              activeTab === "efectuados" 
                ? 'bg-white text-indigo-700 shadow-sm' 
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Pagos Efectuados
            <span className={`ml-2 px-2 py-0.5 rounded-full text-xs ${activeTab === 'efectuados' ? 'bg-indigo-100' : 'bg-slate-200'}`}>
              {pagosEfectuados.length}
            </span>
          </button>
        </div>
      </div>

      {/* Grid AG-Grid */}
      <div className="flex-1 rounded-2xl border border-slate-200 bg-white p-2 shadow-sm overflow-hidden min-h-[500px] flex flex-col">
        <div className="flex-1 w-full relative">
          <div className="absolute inset-0">
            <AgGridReact
              theme={themeQuartz}
              rowData={rowData}
              columnDefs={columnDefs}
              defaultColDef={defaultColDef}
              rowClassRules={rowClassRules}
              pagination={true}
              paginationPageSize={50}
              paginationPageSizeSelector={[50, 100, 200, 500]}
              animateRows={true}
            />
          </div>
        </div>
      </div>
    </div>
  );
}
