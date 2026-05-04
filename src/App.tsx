/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Bus, 
  Search, 
  Plus, 
  TrendingUp, 
  History, 
  MapPin, 
  ArrowLeft,
  Clock,
  Download,
  Navigation,
  FileText
} from 'lucide-react';
import { INITIAL_ROUTES } from './constants';
import { Route, PriceReport, RouteWithStats } from './types';
import { cn } from './lib/utils';
import { 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer 
} from 'recharts';
import * as XLSX from 'xlsx';

export default function App() {
  const [routes] = useState<Route[]>(INITIAL_ROUTES);
  const [reports, setReports] = useState<PriceReport[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedRouteId, setSelectedRouteId] = useState<string | null>(null);
  const [isReporting, setIsReporting] = useState(false);
  const [newPrice, setNewPrice] = useState('');
  const [activeTab, setActiveTab] = useState<'all' | 'reported'>('all');
  const [showMobileSearch, setShowMobileSearch] = useState(false);

  // Load from local storage and add mock if empty
  useEffect(() => {
    const savedReports = localStorage.getItem('taxi_reports_v2');
    if (savedReports) {
      try {
        setReports(JSON.parse(savedReports));
      } catch (e) {
        console.error('Failed to load reports', e);
      }
    } else {
      // Add some initial mock reports for the demo
      const mockReports: PriceReport[] = routes.slice(0, 5).map(r => ({
        id: `mock-${r.id}`,
        routeId: r.id,
        price: Math.floor(Math.random() * 20) + 12,
        timestamp: Date.now() - Math.floor(Math.random() * 3600000),
      }));
      setReports(mockReports);
    }
  }, [routes]);

  // Save to local storage
  useEffect(() => {
    if (reports.length > 0) {
      localStorage.setItem('taxi_reports_v2', JSON.stringify(reports));
    }
  }, [reports]);

  const filteredRoutes = useMemo(() => {
    const baseList = routes.filter(r => 
      r.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.shortCode.toLowerCase().includes(searchQuery.toLowerCase())
    );

    if (activeTab === 'reported') {
      const reportedIds = new Set(reports.map(r => r.routeId));
      return baseList.filter(r => reportedIds.has(r.id));
    }

    return baseList;
  }, [routes, searchQuery, activeTab, reports]);

  const selectedRoute = useMemo(() => {
    if (!selectedRouteId) return null;
    const route = routes.find(r => r.id === selectedRouteId);
    if (!route) return null;

    const routeReports = reports
      .filter(r => r.routeId === selectedRouteId)
      .sort((a, b) => b.timestamp - a.timestamp);

    return {
      ...route,
      currentPrice: routeReports[0]?.price,
      lastUpdated: routeReports[0]?.timestamp,
      history: routeReports
    } as RouteWithStats;
  }, [selectedRouteId, routes, reports]);

  const handleReportPrice = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedRouteId || !newPrice) return;

    const price = parseFloat(newPrice);
    if (isNaN(price)) return;

    const newReport: PriceReport = {
      id: Math.random().toString(36).substr(2, 9),
      routeId: selectedRouteId,
      price,
      timestamp: Date.now(),
    };

    setReports(prev => [newReport, ...prev]);
    setNewPrice('');
    setIsReporting(false);
  };

  const chartData = useMemo(() => {
    if (!selectedRoute) return [];
    return [...selectedRoute.history]
      .reverse()
      .map(r => ({
        time: new Date(r.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        price: r.price
      }));
  }, [selectedRoute]);

  const exportToExcel = async () => {
    let reportsToExport = reports;
    
    // Only export filtered/reported routes if we are on the reported tab
    if (activeTab === 'reported') {
      const reportedRouteIds = new Set(filteredRoutes.map(r => r.id));
      reportsToExport = reports.filter(r => reportedRouteIds.has(r.routeId));
    }

    const data = reportsToExport.map(r => {
      const route = routes.find(rt => rt.id === r.routeId);
      return {
        Route: route?.name || 'Unknown',
        Code: route?.shortCode || '',
        'Price (ETB)': r.price,
        'Date': new Date(r.timestamp).toLocaleDateString(),
        'Time': new Date(r.timestamp).toLocaleTimeString(),
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Taxi Prices");
    
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const fileName = `Addis_Taxi_${activeTab === 'reported' ? 'Reports' : 'All_Prices'}_${new Date().toISOString().split('T')[0]}.xlsx`;

    if (navigator.share && navigator.canShare && navigator.canShare({ files: [new File([blob], fileName, { type: blob.type })] })) {
      try {
        const file = new File([blob], fileName, { type: blob.type });
        await navigator.share({
          title: 'Addis Taxi Prices',
          text: 'Here are the latest minibus taxi prices for Addis Ababa routes!',
          files: [file]
        });
        return;
      } catch (err) {
        console.log('Sharing failed', err);
      }
    }

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-[#FAFAF8] text-[#1A1A1A] font-sans selection:bg-[#FFD700] selection:text-black">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-white/80 backdrop-blur-md border-b border-[#E5E5E1] px-4 py-4 md:px-8">
        <div className="max-w-4xl mx-auto">
          <div className="flex items-center justify-between gap-4">
            <div className={cn("flex items-center gap-3 shrink-0", showMobileSearch && "hidden sm:flex")}>
              <div className="w-10 h-10 bg-[#009739] rounded-xl flex items-center justify-center shadow-lg shadow-[#009739]/20 transform -rotate-3">
                <Bus className="text-white w-6 h-6" />
              </div>
              <div className="hidden xs:block">
                <h1 className="text-xl font-bold tracking-tight">Addis Fare</h1>
                <p className="text-[10px] uppercase tracking-widest text-[#8E9299] font-medium font-mono font-bold">Minibus Finder</p>
              </div>
            </div>
            
            <div className={cn(
              "relative group flex-1 max-w-xs transition-all duration-300",
              showMobileSearch ? "block" : "hidden sm:block"
            )}>
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#8E9299] group-focus-within:text-[#009739] transition-colors font-bold" />
              <input 
                type="text"
                placeholder="Search route (e.g. Bole)..."
                autoFocus={showMobileSearch}
                className="w-full bg-[#F5F5F0] border-none rounded-full py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-[#009739]/20 transition-all outline-none"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                onBlur={() => {
                  if (searchQuery === '') setShowMobileSearch(false);
                }}
              />
              {showMobileSearch && (
                <button 
                  onClick={() => {
                    setShowMobileSearch(false);
                    setSearchQuery('');
                  }}
                  className="sm:hidden absolute right-3 top-1/2 -translate-y-1/2 text-xs font-bold text-[#8E9299]"
                >
                  Clear
                </button>
              )}
            </div>

            <div className="flex items-center gap-2 shrink-0">
              <button 
                onClick={exportToExcel}
                className={cn(
                  "items-center gap-2 bg-[#1A1A1A] text-white px-4 py-2 rounded-full text-sm font-bold shadow-lg shadow-black/10 hover:bg-black transition-all active:scale-95",
                  showMobileSearch ? "hidden sm:flex" : "flex"
                )}
              >
                <Download className="w-4 h-4" />
                <span className="hidden sm:inline">Export Excel</span>
              </button>
              <button 
                onClick={() => setShowMobileSearch(!showMobileSearch)}
                className={cn("sm:hidden p-2 rounded-full hover:bg-[#F5F5F0]", showMobileSearch && "text-[#009739]")}
              >
                <Search className="w-6 h-6" />
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 py-8 md:px-8">
        <AnimatePresence mode="wait">
          {!selectedRouteId ? (
            <motion.div 
              key="list"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <h2 className="text-2xl font-bold font-serif italic font-bold">Taxi Routes</h2>
                  <p className="text-xs text-[#8E9299] font-medium">Select a route to view or report prices</p>
                </div>
                <div className="flex bg-[#F5F5F0] p-1 rounded-xl w-fit self-start sm:self-center">
                  <button 
                    onClick={() => setActiveTab('all')}
                    className={cn(
                      "px-4 py-1.5 rounded-lg text-xs font-bold transition-all",
                      activeTab === 'all' ? "bg-white text-[#1A1A1A] shadow-sm" : "text-[#8E9299] hover:text-[#1A1A1A]"
                    )}
                  >
                    All Routes
                  </button>
                  <button 
                    onClick={() => setActiveTab('reported')}
                    className={cn(
                      "px-4 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-2",
                      activeTab === 'reported' ? "bg-white text-[#1A1A1A] shadow-sm" : "text-[#8E9299] hover:text-[#1A1A1A]"
                    )}
                  >
                    Reported
                    <span className="bg-[#009739]/10 text-[#009739] px-1.5 py-0.5 rounded-full text-[10px]">
                      {new Set(reports.map(r => r.routeId)).size}
                    </span>
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {filteredRoutes.length > 0 ? (
                  filteredRoutes.map((route) => {
                    const routeReports = reports.filter(r => r.routeId === route.id);
                    const lastReport = routeReports[0];
                    
                    return (
                      <motion.button
                        key={route.id}
                        whileHover={{ scale: 1.01 }}
                        whileTap={{ scale: 0.98 }}
                        onClick={() => setSelectedRouteId(route.id)}
                        className="flex items-center justify-between p-5 bg-white border border-[#E5E5E1] rounded-2xl shadow-sm hover:shadow-md transition-all text-left overflow-hidden relative group"
                      >
                        <div className="flex gap-4">
                          <div className="w-12 h-12 rounded-xl bg-[#F5F5F0] flex items-center justify-center text-[#009739] group-hover:bg-[#009739] group-hover:text-white transition-colors">
                            <Navigation className="w-6 h-6" />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <h3 className="font-bold text-base leading-tight">{route.name}</h3>
                            </div>
                            <p className="text-[10px] font-mono font-bold text-[#8E9299] uppercase tracking-wider mt-1">{route.shortCode}</p>
                          </div>
                        </div>
                        <div className="text-right shrink-0 ml-4">
                          {lastReport ? (
                            <>
                              <div className="text-lg font-bold text-[#D21034]">{lastReport.price} <span className="text-[10px] font-medium text-[#8E9299] font-bold">ETB</span></div>
                              <div className="text-[10px] text-[#8E9299] flex items-center justify-end gap-1 font-bold mt-1">
                                <Clock className="w-3 h-3" />
                                {new Date(lastReport.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                              </div>
                            </>
                          ) : (
                            <div className="text-xs text-[#8E9299] italic">No data</div>
                          )}
                        </div>
                      </motion.button>
                    )
                  })
                ) : (
                  <div className="col-span-full py-20 text-center">
                    <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-[#F5F5F0] text-[#8E9299] mb-4">
                      {activeTab === 'reported' ? <FileText className="w-8 h-8" /> : <Search className="w-8 h-8" />}
                    </div>
                    <h3 className="text-lg font-bold text-[#1A1A1A]">No routes found</h3>
                    <p className="text-sm text-[#8E9299] max-w-xs mx-auto mt-1">
                      {activeTab === 'reported' 
                        ? "You haven't reported any prices yet. Switch to 'All Routes' to start reporting."
                        : `We couldn't find any routes matching "${searchQuery}".`}
                    </p>
                  </div>
                )}
              </div>
            </motion.div>
          ) : (
            <motion.div 
              key="detail"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8"
            >
              {/* Detail Header */}
              <div className="flex items-center gap-4">
                <button 
                  onClick={() => setSelectedRouteId(null)}
                  className="p-2 -ml-2 rounded-full hover:bg-[#F5F5F0] transition-colors"
                >
                  <ArrowLeft className="w-6 h-6" />
                </button>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <span className="text-[10px] font-mono bg-[#1A1A1A] text-white px-2 py-0.5 rounded font-bold">{selectedRoute?.shortCode}</span>
                  </div>
                  <h2 className="text-3xl font-bold font-serif italic leading-tight">{selectedRoute?.name}</h2>
                </div>
              </div>

              {/* Stats Cards */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-white p-6 rounded-3xl border border-[#E5E5E1] shadow-sm flex flex-col justify-between h-32 relative overflow-hidden">
                  <span className="text-[10px] uppercase tracking-widest text-[#8E9299] font-bold">Current Fare</span>
                  <div className="flex items-baseline gap-1">
                    <span className="text-4xl font-bold text-[#D21034]">{selectedRoute?.currentPrice || '--'}</span>
                    <span className="text-sm font-medium text-[#8E9299] font-bold">ETB</span>
                  </div>
                  <div className="absolute top-4 right-4 w-12 h-12 rounded-full bg-[#D21034]/5 flex items-center justify-center font-bold">
                    <History className="w-6 h-6 text-[#D21034]" />
                  </div>
                </div>

                <div className="bg-[#1A1A1A] p-6 rounded-3xl shadow-lg shadow-black/10 flex flex-col justify-between h-32 relative text-white">
                  <span className="text-[10px] uppercase tracking-widest text-white/50 font-bold">Status</span>
                  <div className="flex items-center gap-2">
                    <TrendingUp className="text-[#00FF00] w-6 h-6" />
                    <span className="text-lg font-medium">Verified</span>
                  </div>
                  <div className="absolute top-4 right-4 text-white/10 font-bold">
                    <TrendingUp className="w-12 h-12" />
                  </div>
                </div>

                <button 
                  onClick={() => setIsReporting(true)}
                  className="bg-[#009739] p-6 rounded-3xl shadow-lg shadow-[#009739]/20 transform transition-transform hover:scale-[1.02] active:scale-[0.98] flex flex-col justify-between h-32 relative group text-left"
                >
                  <span className="text-[10px] uppercase tracking-widest text-white/70 font-bold">Report Price</span>
                  <div className="flex items-center gap-2 text-white">
                    <Plus className="w-6 h-6" />
                    <span className="text-lg font-medium">Enter Fare</span>
                  </div>
                  <div className="absolute top-4 right-4 w-12 h-12 rounded-full bg-white/10 flex items-center justify-center group-hover:scale-110 transition-transform font-bold">
                    <Plus className="text-white w-6 h-6" />
                  </div>
                </button>
              </div>

              {/* Chart */}
              {selectedRoute && selectedRoute.history.length > 1 && (
                <div className="bg-white p-6 rounded-3xl border border-[#E5E5E1] shadow-sm">
                  <div className="flex items-center justify-between mb-6">
                    <h3 className="font-bold text-lg flex items-center gap-2">
                      <TrendingUp className="w-5 h-5 text-[#009739] font-bold" />
                      Price Trend
                    </h3>
                    <span className="text-[10px] text-[#8E9299] font-mono font-bold">Logs</span>
                  </div>
                  <div className="h-64 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={chartData}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#F0F0F0" />
                        <XAxis 
                          dataKey="time" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{fontSize: 10, fill: '#8E9299'}} 
                          dy={10}
                        />
                        <YAxis 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{fontSize: 10, fill: '#8E9299'}} 
                          dx={-10}
                        />
                        <Tooltip 
                          contentStyle={{ 
                            borderRadius: '12px', 
                            border: 'none', 
                            boxShadow: '0 10px 20px rgba(0,0,0,0.1)',
                            backgroundColor: '#1A1A1A',
                            color: '#fff'
                          }}
                          itemStyle={{ color: '#009739', fontSize: '12px', fontWeight: 'bold' as any }}
                          labelStyle={{ display: 'none' }}
                        />
                        <Line 
                          type="monotone" 
                          dataKey="price" 
                          stroke="#009739" 
                          strokeWidth={3} 
                          dot={{ r: 4, fill: '#009739', stroke: '#fff', strokeWidth: 2 }}
                          activeDot={{ r: 6, fill: '#009739', stroke: '#fff', strokeWidth: 2 }}
                        />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              )}

              {/* Recent Logs */}
              <div className="space-y-4">
                <h3 className="font-bold text-lg flex items-center gap-2">
                  <History className="w-5 h-5 text-[#D21034] font-bold" />
                  Recent Reports
                </h3>
                <div className="space-y-3">
                  {selectedRoute?.history.map((report) => (
                    <div 
                      key={report.id}
                      className="bg-white p-4 rounded-2xl border border-[#E5E5E1] flex items-center justify-between"
                    >
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-full bg-[#F5F5F0] flex items-center justify-center font-bold text-sm text-[#1A1A1A]">
                          A
                        </div>
                        <div>
                          <p className="text-sm font-semibold">User Contribution</p>
                          <p className="text-[10px] text-[#8E9299] flex items-center gap-1 font-bold">
                            <Clock className="w-3 h-3" />
                            {new Date(report.timestamp).toLocaleString()}
                          </p>
                        </div>
                      </div>
                      <div className="font-bold text-[#D21034]">{report.price} <span className="text-[10px] text-[#8E9299] font-bold">ETB</span></div>
                    </div>
                  ))}
                  {selectedRoute?.history.length === 0 && (
                    <div className="text-center py-12 text-[#8E9299] bg-[#F5F5F0] rounded-3xl border-2 border-dashed border-[#E5E5E1]">
                      <Bus className="mx-auto w-12 h-12 mb-3 opacity-20" />
                      <p>No reports for this route yet. Be the first!</p>
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Report Modal */}
      <AnimatePresence>
        {isReporting && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 font-bold">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsReporting(false)}
              className="absolute inset-0 bg-[#000]/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="relative w-full max-w-md bg-white rounded-[32px] overflow-hidden shadow-2xl"
            >
              <div className="p-8 text-center">
                <div className="w-16 h-16 bg-[#009739]/10 rounded-2xl flex items-center justify-center mx-auto mb-6">
                  <Plus className="text-[#009739] w-8 h-8" />
                </div>
                <h3 className="text-2xl font-bold font-serif italic mb-2">Report Price</h3>
                <p className="text-[#8E9299] text-sm mb-8">What is the current fare for <br /><span className="text-[#1A1A1A] font-extrabold">{selectedRoute?.name}</span>?</p>
                
                <form onSubmit={handleReportPrice} className="space-y-6">
                  <div className="relative">
                    <input 
                      type="number"
                      placeholder="0.00"
                      autoFocus
                      className="w-full bg-[#F5F5F0] border-none rounded-2xl py-5 px-6 text-3xl font-bold text-center focus:ring-4 focus:ring-[#009739]/20 transition-all outline-none"
                      value={newPrice}
                      onChange={(e) => setNewPrice(e.target.value)}
                    />
                    <div className="mt-4 text-[10px] uppercase tracking-widest text-[#8E9299] font-extrabold">Price in Ethiopian Birr (ETB)</div>
                  </div>

                  <div className="flex gap-3">
                    <button 
                      type="button"
                      onClick={() => setIsReporting(false)}
                      className="flex-1 py-4 px-6 rounded-2xl bg-[#F5F5F0] font-bold hover:bg-[#E5E5E1] transition-colors"
                    >
                      Cancel
                    </button>
                    <button 
                      type="submit"
                      disabled={!newPrice}
                      className="flex-[2] py-4 px-6 rounded-2xl bg-[#009739] text-white font-bold shadow-lg shadow-[#009739]/20 hover:scale-[1.02] transition-transform disabled:opacity-50 disabled:scale-100 font-extrabold"
                    >
                      Submit Report
                    </button>
                  </div>
                </form>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
