import { ProductivityRecord, RankingEntry } from '../types';
import * as XLSX from 'xlsx';

export function parseProductivityData(rawText: string): ProductivityRecord[] {
  const lines = rawText.trim().split('\n');
  const records: ProductivityRecord[] = [];

  for (const line of lines) {
    // Split by tabs or multiple spaces
    const columns = line.split(/\t/);
    
    // Skip header or invalid lines
    if (columns.length < 5 || columns[0].toLowerCase() === 'processo') continue;

    try {
      const record: ProductivityRecord = {
        processo: columns[0]?.trim() || '',
        atividade: columns[1]?.trim() || '',
        operador: columns[2]?.trim() || '',
        qtdOrdens: parseInt(String(columns[3] || '0').replace(/\D/g, '') || '0', 10),
        qtdPecas: parseInt(String(columns[4] || '0').replace(/\D/g, '') || '0', 10),
        qtdLotes: parseInt(String(columns[5] || '0').replace(/\D/g, '') || '0', 10),
        qtdServicos: parseInt(String(columns[6] || '0').replace(/\D/g, '') || '0', 10),
        qtdItens: parseInt(String(columns[7] || '0').replace(/\D/g, '') || '0', 10),
        qtdEnd: parseInt(String(columns[8] || '0').replace(/\D/g, '') || '0', 10),
        data: columns[9]?.trim() || ''
      };

      if (record.operador) {
        records.push(record);
      }
    } catch (e) {
      console.warn('Failed to parse line:', line, e);
    }
  }

  return records;
}

export function parseExcelData(data: any[]): ProductivityRecord[] {
  return data.map(row => {
    // Standardize column names if they are different in Excel
    // We expect the columns to match the record structure or be in order
    const values = Object.values(row);
    return {
      processo: String(values[0] || ''),
      atividade: String(values[1] || ''),
      operador: String(values[2] || ''),
      qtdOrdens: Number(values[3] || 0),
      qtdPecas: Number(values[4] || 0),
      qtdLotes: Number(values[5] || 0),
      qtdServicos: Number(values[6] || 0),
      qtdItens: Number(values[7] || 0),
      qtdEnd: Number(values[8] || 0),
      data: String(values[9] || '')
    };
  }).filter(r => r.operador && r.operador.toLowerCase() !== 'operador');
}

export function aggregateByEmployee(records: ProductivityRecord[]): Record<string, Omit<RankingEntry, 'rank' | 'score'>> {
  const aggregation: Record<string, any> = {};

  records.forEach(rec => {
    if (!aggregation[rec.operador]) {
      aggregation[rec.operador] = {
        name: rec.operador,
        totalOrdens: 0,
        totalPecas: 0,
        totalServicos: 0,
        totalItens: 0,
      };
    }
    
    aggregation[rec.operador].totalOrdens += rec.qtdOrdens;
    aggregation[rec.operador].totalPecas += rec.qtdPecas;
    aggregation[rec.operador].totalServicos += rec.qtdServicos;
    aggregation[rec.operador].totalItens += rec.qtdItens;
  });

  return aggregation;
}
