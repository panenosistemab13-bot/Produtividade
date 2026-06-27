/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  BarChart3, 
  Upload, 
  Users, 
  Trophy, 
  Download, 
  Trash2, 
  FileSpreadsheet,
  ChevronRight,
  Search,
  Filter,
  PlusCircle,
  LayoutDashboard
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import { parseProductivityData, parseExcelData, aggregateByEmployee } from './utils/parser';
import { ProductivityRecord, EmployeeRoleMap, RankingEntry, Role } from './types';
import { cn } from './lib/utils';

// Professional UI Components
const Card = ({ children, className, id }: { children: React.ReactNode, className?: string, id?: string }) => (
  <div id={id} className={cn("bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden", className)}>
    {children}
  </div>
);

const Button = ({ onClick, children, className, variant = 'primary', disabled }: any) => {
  const variants: any = {
    primary: "bg-slate-900 text-white hover:bg-slate-800",
    secondary: "bg-white border border-slate-200 text-slate-700 hover:bg-slate-50",
    danger: "text-slate-400 hover:text-red-600 hover:bg-red-50",
    success: "bg-emerald-600 text-white hover:bg-emerald-700"
  };
  
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={cn(
        "flex items-center justify-center gap-2 px-4 py-2 rounded-md text-sm font-semibold transition-all duration-200 disabled:opacity-50",
        variants[variant],
        className
      )}
    >
      {children}
    </button>
  );
};

export default function App() {
  const [rawData, setRawData] = useState<string>('');
  const [roleMap, setRoleMap] = useState<EmployeeRoleMap>(() => {
    try {
      const savedRoles = localStorage.getItem('3coracoes_role_map');
      if (savedRoles) {
        const parsed = JSON.parse(savedRoles);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          return parsed;
        }
      }
    } catch (e) {
      console.error('Error parsing 3coracoes_role_map:', e);
    }
    return {};
  });
  const [records, setRecords] = useState<ProductivityRecord[]>(() => {
    try {
      const savedRecords = localStorage.getItem('3coracoes_records');
      if (savedRecords) {
        const parsed = JSON.parse(savedRecords);
        if (Array.isArray(parsed)) {
          return parsed;
        }
      }
    } catch (e) {
      console.error('Error parsing 3coracoes_records:', e);
    }
    return [];
  });
  const [selectedRole, setSelectedRole] = useState<Role | 'Todos'>('Todos');
  const [selectedShift, setSelectedShift] = useState<'Todos' | 'Turno A' | 'Turno B' | 'Turno C'>('Todos');
  const [activeTab, setActiveTab] = useState<'ranking' | 'config' | 'import'>('ranking');
  const [searchTerm, setSearchTerm] = useState('');
  const [configRoleFilter, setConfigRoleFilter] = useState<string>('Todos');
  const [configShiftFilter, setConfigShiftFilter] = useState<string>('Todos');
  const [shiftMap, setShiftMap] = useState<{ [employeeName: string]: string }>(() => {
    try {
      const savedShifts = localStorage.getItem('3coracoes_shift_map');
      if (savedShifts) {
        const parsed = JSON.parse(savedShifts);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          return parsed;
        }
      }
    } catch (e) {
      console.error('Error parsing 3coracoes_shift_map:', e);
    }
    return {};
  });
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Save to local storage
  useEffect(() => {
    if (roleMap && typeof roleMap === 'object') {
      localStorage.setItem('3coracoes_role_map', JSON.stringify(roleMap));
    }
  }, [roleMap]);

  useEffect(() => {
    if (shiftMap && typeof shiftMap === 'object') {
      localStorage.setItem('3coracoes_shift_map', JSON.stringify(shiftMap));
    }
  }, [shiftMap]);

  useEffect(() => {
    if (Array.isArray(records)) {
      localStorage.setItem('3coracoes_records', JSON.stringify(records));
    }
  }, [records]);

  const handleImport = () => {
    const newRecords = parseProductivityData(rawData);
    if (newRecords.length > 0) {
      setRecords(prev => [...newRecords]); 
      setRawData('');
      setActiveTab('ranking');
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const newRecords = parseExcelData(data);
      if (newRecords.length > 0) {
        setRecords(prev => [...newRecords]);
        setActiveTab('ranking');
      }
    };
    reader.readAsBinaryString(file);
  };

  const clearData = () => {
    if (confirm('Tem certeza que deseja limpar todos os dados?')) {
      setRecords([]);
      localStorage.removeItem('3coracoes_records');
    }
  };

  const setEmployeeRole = (name: string, role: string) => {
    setRoleMap(prev => ({ ...prev, [name]: role }));
  };

  const setEmployeeShift = (name: string, shift: string) => {
    setShiftMap(prev => ({ ...prev, [name]: shift }));
  };

  // Processing rankings
  const aggregated = aggregateByEmployee(records);
  const employeeNames = Object.keys(aggregated);
  const roles = Array.from(new Set(Object.values(roleMap)));

  const getRankings = (filterRole: string, filterShift: string): RankingEntry[] => {
    let filteredNames = employeeNames;
    if (filterRole !== 'Todos') {
      filteredNames = filteredNames.filter(name => roleMap[name] === filterRole);
    }
    if (filterShift !== 'Todos') {
      filteredNames = filteredNames.filter(name => {
        const shift = shiftMap[name] || '';
        if (filterShift === 'Sem Turno') {
          return !shift;
        }
        return shift === filterShift;
      });
    }

    if (searchTerm) {
      filteredNames = filteredNames.filter(name => name.toLowerCase().includes(searchTerm.toLowerCase()));
    }

    return filteredNames
      .map(name => {
        const data = aggregated[name];
        const score = data.totalServicos + (data.totalOrdens * 0.5) + (data.totalItens * 0.1);
        return {
          ...data,
          score,
          rank: 0
        };
      })
      .sort((a, b) => b.score - a.score)
      .map((entry, index) => ({ ...entry, rank: index + 1 }));
  };

  const rankings = getRankings(selectedRole, selectedShift);

  const downloadRanking = async (role: string, shift: string) => {
    const element = document.getElementById(`ranking-capture-area`);
    if (element) {
      try {
        const dataUrl = await toPng(element, { 
          quality: 1.0, 
          backgroundColor: '#ffffff',
          pixelRatio: 3,
          cacheBust: true
        });
        const link = document.createElement('a');
        const roleStr = role.toLowerCase().replace(/\s+/g, '-');
        const shiftStr = shift.toLowerCase().replace(/\s+/g, '-');
        link.download = `ranking-${roleStr}-${shiftStr}.png`;
        link.href = dataUrl;
        link.click();
      } catch (err) {
        console.error('Download failed', err);
      }
    }
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] text-slate-900 font-sans">
      {/* Top Navigation Bar */}
      <nav className="bg-white border-b border-slate-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-8">
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 bg-slate-900 rounded flex items-center justify-center">
                <BarChart3 className="text-white w-5 h-5" />
              </div>
              <span className="font-bold text-lg tracking-tight">Logística 3 Corações</span>
            </div>
            
            <div className="hidden md:flex items-center gap-1">
              {[
                { id: 'ranking', label: 'Painel de Ranking', icon: LayoutDashboard },
                { id: 'import', label: 'Importação', icon: Upload },
                { id: 'config', label: 'Gestão de Funções', icon: Users },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id as any)}
                  className={cn(
                    "flex items-center gap-2 px-4 py-2 rounded-md transition-all text-sm font-bold",
                    activeTab === tab.id 
                      ? "bg-slate-100 text-slate-900" 
                      : "text-slate-500 hover:bg-slate-50 hover:text-slate-900"
                  )}
                >
                  <tab.icon className="w-4 h-4" />
                  {tab.label}
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right hidden lg:block">
              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-tight">criado por: Jefferson Augusto (10-85447)</p>
              <p className="text-[9px] font-bold text-slate-900 uppercase tracking-widest leading-tight">abra o formulario na tela: 806</p>
            </div>
          </div>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 py-8">
        <AnimatePresence mode="wait">
          {activeTab === 'import' && (
            <motion.div
              key="import"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="max-w-4xl mx-auto"
            >
              <div className="mb-6 flex items-center justify-between">
                <div>
                  <h1 className="text-2xl font-bold">Importação de Dados</h1>
                  <p className="text-slate-500 text-sm">Carregue arquivos Excel ou cole dados de texto para atualizar o sistema.</p>
                </div>
                <div className="flex gap-2">
                  <input 
                    type="file" 
                    ref={fileInputRef}
                    onChange={handleFileUpload}
                    accept=".xlsx, .xls, .csv"
                    className="hidden"
                  />
                  <Button variant="secondary" onClick={() => fileInputRef.current?.click()}>
                    <FileSpreadsheet className="w-4 h-4" />
                    Upload Excel
                  </Button>
                  <Button variant="primary" onClick={handleImport} disabled={!rawData.trim()}>
                    Processar Texto
                  </Button>
                </div>
              </div>

              <Card className="p-0 border-slate-200">
                <textarea
                  value={rawData}
                  onChange={(e) => setRawData(e.target.value)}
                  placeholder="Cole os dados da planilha aqui..."
                  className="w-full h-96 p-6 text-sm font-mono focus:outline-none resize-none bg-slate-50/30"
                />
              </Card>

              <div className="mt-4 flex justify-end">
                <Button variant="danger" onClick={clearData}>
                  <Trash2 className="w-4 h-4" />
                  Zerar Todos os Dados
                </Button>
              </div>
            </motion.div>
          )}

          {activeTab === 'config' && (
            <motion.div
              key="config"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="max-w-4xl mx-auto"
            >
              <div className="mb-6 flex items-center justify-between">
                <div>
                  <h1 className="text-2xl font-bold">Gestão de Colaboradores</h1>
                  <p className="text-slate-500 text-sm">Atribua funções e turnos específicos para segmentar os rankings mensais.</p>
                </div>
                <div className="flex flex-wrap items-center gap-3">
                  <div className="relative">
                    <Filter className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <select
                      value={configShiftFilter}
                      onChange={(e) => setConfigShiftFilter(e.target.value)}
                      className="bg-white border border-slate-200 rounded-md pl-10 pr-4 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-slate-400 w-44 appearance-none"
                    >
                      <option value="Todos">Todos Turnos</option>
                      <option value="Sem Turno">Sem Turno</option>
                      <option value="Turno A">Turno A</option>
                      <option value="Turno B">Turno B</option>
                      <option value="Turno C">Turno C</option>
                    </select>
                  </div>
                  <div className="relative">
                    <Filter className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <select
                      value={configRoleFilter}
                      onChange={(e) => setConfigRoleFilter(e.target.value)}
                      className="bg-white border border-slate-200 rounded-md pl-10 pr-4 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-slate-400 w-44 appearance-none"
                    >
                      <option value="Todos">Todas Funções</option>
                      <option value="Sem Função">Sem Função</option>
                      <option value="Auxiliar">Auxiliar</option>
                      <option value="Operador">Operador</option>
                      <option value="Conferente">Conferente</option>
                      <option value="Lider">Líder</option>
                      <option value="Supervisor">Supervisor</option>
                    </select>
                  </div>
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input 
                      type="text"
                      placeholder="Filtrar por nome..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="bg-white border border-slate-200 rounded-md pl-10 pr-4 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-slate-400 w-64"
                    />
                  </div>
                </div>
              </div>

              <Card className="border-slate-200">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>
                        <th className="px-6 py-3 font-black text-slate-900">Colaborador</th>
                        <th className="px-6 py-3 font-black text-slate-900 text-right">Turno</th>
                        <th className="px-6 py-3 font-black text-slate-900 text-right">Função Designada</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {employeeNames
                        .filter(name => {
                          const matchesSearch = name.toLowerCase().includes(searchTerm.toLowerCase());
                          const role = roleMap[name] || 'Sem Função';
                          const shift = shiftMap[name] || '';
                          const matchesRole = configRoleFilter === 'Todos' || role === configRoleFilter;
                          const matchesShift = configShiftFilter === 'Todos' || 
                            (configShiftFilter === 'Sem Turno' && !shift) || 
                            shift === configShiftFilter;
                          return matchesSearch && matchesRole && matchesShift;
                        })
                        .map(name => (
                          <tr key={name} className="hover:bg-slate-50 transition-colors">
                            <td className="px-6 py-4 font-black text-slate-900">{name}</td>
                            <td className="px-6 py-4 text-right">
                              <select
                                value={shiftMap[name] || ''}
                                onChange={(e) => setEmployeeShift(name, e.target.value)}
                                className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-slate-400"
                              >
                                <option value="">Sem Turno</option>
                                <option value="Turno A">Turno A</option>
                                <option value="Turno B">Turno B</option>
                                <option value="Turno C">Turno C</option>
                              </select>
                            </td>
                            <td className="px-6 py-4 text-right">
                              <select
                                value={roleMap[name] || ''}
                                onChange={(e) => setEmployeeRole(name, e.target.value)}
                                className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-slate-400"
                              >
                                <option value="">Não Atribuída</option>
                                <option value="Auxiliar">Auxiliar</option>
                                <option value="Operador">Operador</option>
                                <option value="Conferente">Conferente</option>
                                <option value="Lider">Líder</option>
                                <option value="Supervisor">Supervisor</option>
                              </select>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            </motion.div>
          )}

          {activeTab === 'ranking' && (
            <motion.div
              key="ranking"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="space-y-6"
            >
              {/* Header & Controls */}
              <div className="flex flex-col lg:flex-row gap-6 items-start lg:items-center justify-between">
                <div>
                  <h1 className="text-3xl font-bold tracking-tight">Ranking de Produtividade</h1>
                  <p className="text-slate-500 text-sm font-medium">Relatório mensal de desempenho operacional</p>
                </div>
                
                <div className="flex flex-wrap items-center gap-4">
                  {/* Filtro por Turno */}
                  <div className="flex flex-col gap-1">
                    <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">Filtrar por Turno</span>
                    <div className="flex bg-white border border-slate-200 p-1 rounded-md">
                      {['Todos', 'Turno A', 'Turno B', 'Turno C'].map(shift => (
                        <button
                          key={shift}
                          onClick={() => setSelectedShift(shift as any)}
                          className={cn(
                            "px-3 py-1.5 rounded text-xs font-bold transition-all",
                            selectedShift === shift 
                              ? "bg-slate-900 text-white shadow-sm" 
                              : "text-slate-500 hover:text-slate-900"
                          )}
                        >
                          {shift}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Filtro por Função */}
                  <div className="flex flex-col gap-1">
                    <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">Filtrar por Função</span>
                    <div className="flex bg-white border border-slate-200 p-1 rounded-md">
                      {['Todos', ...roles].filter(Boolean).map(role => (
                        <button
                          key={role}
                          onClick={() => setSelectedRole(role)}
                          className={cn(
                            "px-3 py-1.5 rounded text-xs font-bold transition-all",
                            selectedRole === role 
                              ? "bg-slate-900 text-white shadow-sm" 
                              : "text-slate-500 hover:text-slate-900"
                          )}
                        >
                          {role}
                        </button>
                      ))}
                    </div>
                  </div>
                  
                  <div className="flex flex-col gap-1 self-end">
                    <span className="text-[10px] font-black uppercase text-transparent select-none tracking-wider">Exportar</span>
                    <Button variant="success" onClick={() => downloadRanking(selectedRole, selectedShift)}>
                      <Download className="w-4 h-4" />
                      Exportar Ranking
                    </Button>
                  </div>
                </div>
              </div>

              {/* Data Table */}
              <div id="ranking-capture-area">
                <Card className="border-slate-300 shadow-lg">
                  {/* Internal Branding for Export */}
                  <div className="bg-slate-900 p-6 flex justify-between items-center">
                    <div>
                      <h2 className="text-white text-xl font-bold flex items-center gap-2">
                        <Trophy className="text-amber-400 w-5 h-5" />
                        Top Performance: {selectedRole === 'Todos' ? 'Geral' : selectedRole} {selectedShift !== 'Todos' ? ` • ${selectedShift}` : ''}
                      </h2>
                      <p className="text-slate-400 text-xs font-medium uppercase tracking-widest mt-1">3 Corações Logistics • Junho 2026</p>
                    </div>
                    <div className="text-right">
                      <div className="text-white font-bold text-xs">criado por: Jefferson Augusto (10-85447)</div>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-slate-50 border-b border-slate-200">
                        <tr>
                          <th className="px-6 py-4 font-black text-slate-900 w-16">Pos</th>
                          <th className="px-6 py-4 font-black text-slate-900">Colaborador</th>
                          <th className="px-6 py-4 font-black text-slate-900 text-center">Serviços</th>
                          <th className="px-6 py-4 font-black text-slate-900 text-center">Ordens</th>
                          <th className="px-6 py-4 font-black text-slate-900 text-center">Itens</th>
                          <th className="px-6 py-4 font-black text-slate-900 text-right">Pontuação</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {rankings.length > 0 ? (
                          rankings.map((entry) => (
                            <tr key={entry.name} className={cn(
                              "hover:bg-slate-50 transition-colors",
                              entry.rank === 1 ? "bg-slate-50/50" : ""
                            )}>
                              <td className="px-6 py-4 font-black">
                                <div className={cn(
                                  "w-8 h-8 rounded-full flex items-center justify-center text-xs",
                                  entry.rank === 1 ? "bg-amber-100 text-amber-700 border border-amber-200" :
                                  entry.rank === 2 ? "bg-slate-100 text-slate-700 border border-slate-200" :
                                  entry.rank === 3 ? "bg-orange-50 text-orange-700 border border-orange-100" :
                                  "text-slate-900"
                                )}>
                                  {entry.rank}
                                </div>
                              </td>
                              <td className="px-6 py-4">
                                <span className="font-black text-slate-900">{entry.name}</span>
                                {entry.rank === 1 && <span className="ml-2 text-[10px] bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded font-black uppercase tracking-tighter">Líder do Ranking</span>}
                              </td>
                              <td className="px-6 py-4 text-center font-mono font-bold text-slate-900">{entry.totalServicos}</td>
                              <td className="px-6 py-4 text-center font-mono font-bold text-slate-900">{entry.totalOrdens}</td>
                              <td className="px-6 py-4 text-center font-mono font-bold text-slate-900">{entry.totalItens}</td>
                              <td className="px-6 py-4 text-right">
                                <span className="text-lg font-black text-slate-900">{Math.round(entry.score)}</span>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={6} className="px-6 py-24 text-center text-slate-400 font-medium italic">
                              Nenhum dado disponível para este filtro.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                  
                  {/* Footer Decoration for Export */}
                  <div className="bg-slate-50 px-6 py-3 border-t border-slate-200 flex justify-between items-center">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest italic">Confidencial • Uso Interno</span>
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{new Date().toLocaleDateString('pt-BR')}</span>
                  </div>
                </Card>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Corporate Footer */}
      <footer className="max-w-7xl mx-auto px-4 py-12 border-t border-slate-200 mt-12 flex justify-between items-center text-slate-400">
        <div className="flex items-center gap-2">
          <div className="w-6 h-6 bg-slate-200 rounded flex items-center justify-center">
            <BarChart3 className="text-slate-400 w-3 h-3" />
          </div>
          <span className="text-[10px] font-bold uppercase tracking-[0.2em]">Sistema Jefferson Augusto</span>
        </div>
        <p className="text-[10px] font-bold uppercase tracking-[0.2em]">© 2026 Grupo 3 Corações Logística</p>
      </footer>
    </div>
  );
}
