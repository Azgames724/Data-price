/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface Route {
  id: string;
  name: string; // Long name: "Kality ↔ Bole"
  shortCode: string; // TX001
  description: string;
}

export interface PriceReport {
  id: string;
  routeId: string;
  price: number; // in ETB
  timestamp: number;
  reporter?: string;
}

export interface RouteWithStats extends Route {
  currentPrice?: number;
  lastUpdated?: number;
  history: PriceReport[];
}
