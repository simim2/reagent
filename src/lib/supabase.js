import { createClient } from '@supabase/supabase-js'

export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY,
)

// ── DB row ↔ app object 변환 ────────────────────────────────
export const toAppReagent = (r) => ({
  id: r.id,
  name: r.name,
  reagentType: r.reagent_type ?? 'Reagent',
  manufacturer: r.manufacturer,
  lotNo: r.lot_no,
  receivedQty: r.received_qty,
  currentStock: r.current_stock,
  minStock: r.min_stock,
  expiryDate: r.expiry_date,
  totalDispatched: r.total_dispatched,
  createdAt: r.created_at ?? null,
})

export const toDbReagent = (r) => ({
  id: r.id,
  name: r.name,
  reagent_type: r.reagentType ?? 'Reagent',
  manufacturer: r.manufacturer,
  lot_no: r.lotNo,
  received_qty: r.receivedQty,
  current_stock: r.currentStock,
  min_stock: r.minStock,
  expiry_date: r.expiryDate,
  total_dispatched: r.totalDispatched,
  updated_at: new Date().toISOString(),
})

export const toAppLog = (l) => ({
  id: l.id,
  reagentId: l.reagent_id,
  reagentName: l.reagent_name,
  lotNo: l.lot_no,
  qty: l.qty,
  datetime: l.datetime,
})

export const toDbLog = (l) => ({
  id: l.id,
  reagent_id: l.reagentId,
  reagent_name: l.reagentName,
  lot_no: l.lotNo,
  qty: l.qty,
  datetime: l.datetime,
})

export const toAppInbound = (l) => ({
  id: l.id,
  reagentId: l.reagent_id,
  reagentName: l.reagent_name,
  lotNo: l.lot_no,
  qty: l.qty,
  datetime: l.datetime,
  notes: l.notes ?? '',
})

export const toDbInbound = (l) => ({
  id: l.id,
  reagent_id: l.reagentId,
  reagent_name: l.reagentName,
  lot_no: l.lotNo,
  qty: l.qty,
  datetime: l.datetime,
  notes: l.notes ?? '',
})
