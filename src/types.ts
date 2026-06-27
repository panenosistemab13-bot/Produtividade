/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface ProductivityRecord {
  processo: string;
  atividade: string;
  operador: string;
  qtdOrdens: number;
  qtdPecas: number;
  qtdLotes: number;
  qtdServicos: number;
  qtdItens: number;
  qtdEnd: number;
  data: string;
}

export type Role = string;

export interface EmployeeRoleMap {
  [employeeName: string]: Role;
}

export interface RankingEntry {
  name: string;
  totalOrdens: number;
  totalPecas: number;
  totalServicos: number;
  totalItens: number;
  score: number;
  rank: number;
}
