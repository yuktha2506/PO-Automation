import React, { useState, useRef, useEffect } from 'react';
import { POData } from '../types';
import { Edit2, Check, ChevronDown, ChevronRight, AlertTriangle, X, Hash, Calendar, User, Tag } from 'lucide-react';

interface DataPreviewProps {
  data: POData[];
  onUpdate: (id: string, updatedData: Partial<POData>) => void;
}

const TAX_RATE = 0.18;

const toSafeNumber = (value: unknown): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
};

const calculateLineTotals = (quantity: unknown, unitRate: unknown) => {
  const baseAmount = toSafeNumber(quantity) * toSafeNumber(unitRate);
  const tax = baseAmount * TAX_RATE;
  const grandTotal = baseAmount + tax;
  return { baseAmount, tax, grandTotal };
};

export const DataPreview: React.FC<DataPreviewProps> = ({ data, onUpdate }) => {
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<POData | null>(null);
  const editButtonRefs = useRef<{ [key: string]: HTMLButtonElement | null }>({});

  const toggleExpand = (id: string) => {
    setExpandedId(expandedId === id ? null : id);
  };

  const startEditing = (po: POData) => {
    setEditingId(po.id);
    setEditForm({ ...po });
  };

  const saveEdit = () => {
    if (editForm && editingId) {
      onUpdate(editingId, editForm);
      setEditingId(null);
      setEditForm(null);
    }
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditForm(null);
  };

  const handleInputChange = (field: keyof POData, value: string | number) => {
    if (editForm) {
      setEditForm({ ...editForm, [field]: value });
    }
  };

  if (data.length === 0) return null;

  return (
    <div className="bg-white/70 dark:bg-slate-800/70 backdrop-blur-xl rounded-3xl border border-white/40 dark:border-slate-700 shadow-xl overflow-hidden flex flex-col max-h-[800px]">
      <div className="px-8 py-5 border-b border-gray-200/60 dark:border-slate-700/60 flex justify-between items-center bg-white/40 dark:bg-slate-800/40">
        <div>
          <h3 className="text-xl font-bold text-gray-800 dark:text-gray-100">Extracted Data</h3>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">Reviewing {data.length} records</p>
        </div>
      </div>

      <div className="overflow-auto flex-grow scrollbar-thin scrollbar-thumb-gray-300 dark:scrollbar-thumb-slate-600">
        <table className="w-full text-xs text-left">
          <thead className="text-xs font-bold text-gray-500 dark:text-gray-400 uppercase bg-gray-50/80 dark:bg-slate-900/60 border-b border-gray-200 dark:border-slate-700 sticky top-0 z-10 backdrop-blur-md">
            <tr>
              <th className="px-4 py-3 w-10 text-center">#</th>
              <th className="px-4 py-3">Category</th>
              <th className="px-4 py-3">Requestor</th>
              <th className="px-4 py-3">PR Info</th>
              <th className="px-4 py-3">PO Info</th>
              <th className="px-4 py-3">Vendor</th>
              <th className="px-4 py-3 text-right">Total</th>
              <th className="px-4 py-3 text-center">Status</th>
              <th className="px-4 py-3 text-center">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100 dark:divide-slate-700/50">
            {data.map((po) => {
              const isEditing = editingId === po.id;
              
              return (
                <React.Fragment key={po.id}>
                  <tr className={`${isEditing ? 'bg-blue-50/60 dark:bg-blue-900/20' : 'hover:bg-blue-50/30 dark:hover:bg-slate-700/30'}`}>
                    <td className="px-4 py-4 text-center">
                      <button 
                        onClick={() => toggleExpand(po.id)}
                        className="p-1 rounded hover:bg-gray-200 dark:hover:bg-slate-600 transition-colors"
                      >
                         {expandedId === po.id ? <ChevronDown className="w-3 h-3"/> : <ChevronRight className="w-3 h-3"/>}
                      </button>
                    </td>

                    {isEditing && editForm ? (
                      <>
                        <td className="px-4 py-3 align-top">
                          <input 
                            type="text" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.category}
                            onChange={(e) => handleInputChange('category', e.target.value)}
                            placeholder="Category"
                          />
                        </td>
                        <td className="px-4 py-3 align-top">
                          <input 
                            type="text" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.requestor_name}
                            onChange={(e) => handleInputChange('requestor_name', e.target.value)}
                            placeholder="Requestor"
                          />
                        </td>
                         <td className="px-4 py-3 align-top space-y-1">
                          <input 
                            type="text" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.pr_number}
                            onChange={(e) => handleInputChange('pr_number', e.target.value)}
                            placeholder="PR #"
                          />
                          <input 
                            type="date" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.pr_date}
                            onChange={(e) => handleInputChange('pr_date', e.target.value)}
                          />
                        </td>
                        <td className="px-4 py-3 align-top space-y-1">
                          <input 
                            type="text" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.po_number}
                            onChange={(e) => handleInputChange('po_number', e.target.value)}
                            placeholder="PO #"
                          />
                          <input 
                            type="date" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.date}
                            onChange={(e) => handleInputChange('date', e.target.value)}
                          />
                        </td>
                        <td className="px-4 py-3 align-top">
                          <input 
                            type="text" 
                            className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                            value={editForm.supplier_name}
                            onChange={(e) => handleInputChange('supplier_name', e.target.value)}
                          />
                        </td>
                        <td className="px-4 py-3 align-top text-right">
                          <input 
                            type="number" 
                            className="w-20 px-2 py-1 text-xs text-right border rounded bg-white dark:bg-slate-800 dark:border-slate-600 font-mono"
                            value={editForm.total}
                            onChange={(e) => handleInputChange('total', parseFloat(e.target.value))}
                          />
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="px-4 py-4 align-top">
                          <span className="inline-block px-2 py-0.5 rounded text-[10px] bg-purple-100 text-purple-800 dark:bg-purple-900/50 dark:text-purple-300 font-semibold">{po.category}</span>
                        </td>
                        <td className="px-4 py-4 align-top font-medium text-gray-700 dark:text-gray-300">{po.requestor_name}</td>
                        <td className="px-4 py-4 align-top">
                          <div className="font-medium text-gray-900 dark:text-gray-100">{po.pr_number}</div>
                          <div className="text-[10px] text-gray-400 font-mono">{po.pr_date}</div>
                        </td>
                        <td className="px-4 py-4 align-top">
                          <div className="font-bold text-blue-600 dark:text-blue-400">{po.po_number}</div>
                          <div className="text-[10px] text-gray-400 font-mono">{po.date}</div>
                        </td>
                        <td className="px-4 py-4 align-top text-gray-600 dark:text-gray-400 truncate max-w-[150px]" title={po.supplier_name}>
                          {po.supplier_name}
                        </td>
                        <td className="px-4 py-4 align-top text-right font-mono font-bold text-gray-900 dark:text-white">
                          {po.total.toFixed(2)}
                        </td>
                      </>
                    )}

                    <td className="px-4 py-4 align-top text-center">
                       {po.status === 'completed' ? (
                        <span className="text-xs text-emerald-600 dark:text-emerald-400 font-semibold">Verified</span>
                      ) : po.status === 'error' ? (
                        <span className="text-xs text-red-600 dark:text-red-400 font-bold" title={po.errorMessage}>Error</span>
                      ) : (
                         <span className="text-xs text-amber-600 dark:text-amber-400">Pending</span>
                      )}
                    </td>
                    <td className="px-4 py-4 align-top text-center">
                       {isEditing ? (
                         <div className="flex space-x-1 justify-center">
                           <button onClick={saveEdit} className="p-1.5 bg-emerald-100 text-emerald-600 rounded hover:bg-emerald-200"><Check className="w-3 h-3"/></button>
                           <button onClick={cancelEdit} className="p-1.5 bg-red-100 text-red-600 rounded hover:bg-red-200"><X className="w-3 h-3"/></button>
                         </div>
                       ) : (
                         <button onClick={() => startEditing(po)} disabled={po.status !== 'completed'} className="p-1.5 text-blue-600 hover:bg-blue-50 rounded disabled:opacity-30"><Edit2 className="w-3 h-3"/></button>
                       )}
                    </td>
                  </tr>

                  {/* Expanded Row for Line Items & Additional Fields */}
                  {expandedId === po.id && (
                    <tr className="bg-gray-50 dark:bg-slate-800/50">
                      <td colSpan={9} className="p-4">
                        <div className="bg-white dark:bg-slate-900 rounded-lg border border-gray-200 dark:border-slate-700 overflow-hidden space-y-4 p-4">
                          
                          {/* Additional Fields for Tracker */}
                          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pb-4 border-b border-gray-100 dark:border-slate-800">
                             <div>
                               <label className="text-[10px] text-gray-400 uppercase font-bold block mb-1">Delivery Date (Agreed)</label>
                               {isEditing && editForm ? (
                                  <input 
                                    type="date" 
                                    className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                                    value={editForm.delivery_date_agreed || ''}
                                    onChange={(e) => handleInputChange('delivery_date_agreed', e.target.value)}
                                  />
                               ) : (
                                  <div className="text-xs font-mono">{po.delivery_date_agreed || '-'}</div>
                               )}
                             </div>
                             <div>
                               <label className="text-[10px] text-gray-400 uppercase font-bold block mb-1">Remarks</label>
                               {isEditing && editForm ? (
                                  <input 
                                    type="text" 
                                    className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                                    value={editForm.remarks || ''}
                                    onChange={(e) => handleInputChange('remarks', e.target.value)}
                                  />
                               ) : (
                                  <div className="text-xs">{po.remarks || '-'}</div>
                               )}
                             </div>
                             <div>
                               <label className="text-[10px] text-gray-400 uppercase font-bold block mb-1">No of Days (Remarks)</label>
                               {isEditing && editForm ? (
                                  <input 
                                    type="text" 
                                    className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                                    value={editForm.no_of_days_remarks || ''}
                                    onChange={(e) => handleInputChange('no_of_days_remarks', e.target.value)}
                                  />
                               ) : (
                                  <div className="text-xs">{po.no_of_days_remarks || '-'}</div>
                               )}
                             </div>
                             <div>
                               <label className="text-[10px] text-gray-400 uppercase font-bold block mb-1">Negotiation Remarks</label>
                               {isEditing && editForm ? (
                                  <input 
                                    type="text" 
                                    className="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-800 dark:border-slate-600"
                                    value={editForm.negotiation_remarks || ''}
                                    onChange={(e) => handleInputChange('negotiation_remarks', e.target.value)}
                                  />
                               ) : (
                                  <div className="text-xs">{po.negotiation_remarks || '-'}</div>
                               )}
                             </div>
                          </div>

                          <div>
                            <div className="text-xs font-bold text-gray-500 uppercase mb-2">Item Details</div>
                            <table className="w-full text-xs">
                              <thead className="border-b border-gray-200 dark:border-slate-700 bg-gray-50 dark:bg-slate-800/50">
                                <tr>
                                  <th className="px-4 py-2 text-left">Description</th>
                                  <th className="px-4 py-2 text-right">Qty</th>
                                  <th className="px-4 py-2 text-right">Rate</th>
                                  <th className="px-4 py-2 text-right">Amount</th>
                                  <th className="px-4 py-2 text-right">Tax</th>
                                  <th className="px-4 py-2 text-right">Grand Total</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-gray-100 dark:divide-slate-800">
                                 {po.items.map((item, idx) => {
                                   const totals = calculateLineTotals(item.quantity, item.rate);
                                   return (
                                     <tr key={idx}>
                                       <td className="px-4 py-2">{item.description}</td>
                                       <td className="px-4 py-2 text-right font-mono">{toSafeNumber(item.quantity)}</td>
                                       <td className="px-4 py-2 text-right font-mono">{toSafeNumber(item.rate).toFixed(2)}</td>
                                       <td className="px-4 py-2 text-right font-mono font-semibold">{totals.baseAmount.toFixed(2)}</td>
                                       <td className="px-4 py-2 text-right font-mono">{totals.tax.toFixed(2)}</td>
                                       <td className="px-4 py-2 text-right font-mono font-semibold">{totals.grandTotal.toFixed(2)}</td>
                                     </tr>
                                   );
                                 })}
                                 <tr className="bg-gray-50/50">
                                   <td colSpan={5} className="px-4 py-2 text-right font-medium">Combined Grand Total</td>
                                   <td className="px-4 py-2 text-right font-mono font-semibold">
                                     {po.items.reduce((sum, item) => sum + calculateLineTotals(item.quantity, item.rate).grandTotal, 0).toFixed(2)}
                                   </td>
                                 </tr>
                              </tbody>
                            </table>
                          </div>
                          
                          <div className="p-3 border-t border-gray-100 dark:border-slate-800 text-xs text-gray-500">
                             <span className="font-semibold">Description Summary:</span> {po.description}
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};
