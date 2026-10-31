import { BusinessDocument, BusinessSettings, Client, Product } from '@/types';

const keys = { clients: 'fg_clients', products: 'fg_products', documents: 'fg_documents', settings: 'fg_settings' } as const;
const starterClients: Client[] = [
  { id: 'c1', name: 'Constructora del Valle', attention: 'María López', email: 'compras@constructoravalle.com', phone: '55 1234 5678' },
  { id: 'c2', name: 'Retama', attention: 'Jorge Ramírez', email: 'jorge@retama.mx', phone: '55 4589 2210' },
  { id: 'c3', name: 'Obras García', attention: 'Luis García', email: 'luis@obrasgarcia.mx', phone: '55 6812 9001' },
  { id: 'c4', name: 'Mantenimiento Norte', attention: 'Ana Torres', email: 'ana@mantenimientonorte.mx', phone: '55 3011 7482' },
];
const starterProducts: Product[] = [
  { id: 'p1', name: 'Cemento gris 50 kg', unit: 'bulto', price: 285 }, { id: 'p2', name: 'Varilla corrugada 3/8', unit: 'pza', price: 145 },
  { id: 'p3', name: 'Arena fina', unit: 'bulto', price: 68 }, { id: 'p4', name: 'Grava 3/4', unit: 'bulto', price: 72 },
  { id: 'p5', name: 'Pegazulejo blanco', unit: 'bulto', price: 198 }, { id: 'p6', name: 'Tubo PVC hidráulico 1/2', unit: 'pza', price: 39 },
  { id: 'p7', name: 'Llave de paso', unit: 'pza', price: 119 }, { id: 'p8', name: 'Pintura vinílica blanca', unit: 'lt', price: 145 },
  { id: 'p9', name: 'Brocha 4 pulgadas', unit: 'pza', price: 84 }, { id: 'p10', name: 'Rodillo profesional', unit: 'pza', price: 135 },
  { id: 'p11', name: 'Cinta teflón', unit: 'pza', price: 24 }, { id: 'p12', name: 'Disco de corte', unit: 'pza', price: 35 },
  { id: 'p13', name: 'Taquete expansivo', unit: 'caja', price: 115 }, { id: 'p14', name: 'Silicón transparente', unit: 'pza', price: 78 },
  { id: 'p15', name: 'Guantes de carnaza', unit: 'pza', price: 62 },
];
const starterDocuments: BusinessDocument[] = [
  { id: 'd-r1171', type: 'remision', number: 1171, date: '2026-10-22', clientId: 'c1', clientName: 'Constructora del Valle', attention: 'María López', email: 'compras@constructoravalle.com', phone: '55 1234 5678', items: [{ id: 'i-r1171', quantity: 6, unit: 'pza', description: 'Botes de poliuretano', unitPrice: 345 }], delivery: 'Inmediata', payment: 'Por definir', note: '', createdAt: '2026-10-22T12:00:00.000Z' },
  { id: 'd-r1170', type: 'remision', number: 1170, date: '2026-10-21', clientId: 'c2', clientName: 'Retama', attention: 'Jorge Ramírez', email: 'jorge@retama.mx', phone: '55 4589 2210', items: [{ id: 'i-r1170', quantity: 5, unit: 'pza', description: 'Silicón negro', unitPrice: 125 }], delivery: 'Inmediata', payment: 'Por definir', note: '', createdAt: '2026-10-21T12:00:00.000Z' },
  { id: 'd-r1169', type: 'remision', number: 1169, date: '2026-10-20', clientId: 'c3', clientName: 'Obras García', attention: 'Luis García', email: 'luis@obrasgarcia.mx', phone: '55 6812 9001', items: [{ id: 'i-r1169', quantity: 5, unit: 'pza', description: 'Silicón blanco', unitPrice: 125 }], delivery: 'Inmediata', payment: 'Por definir', note: '', createdAt: '2026-10-20T12:00:00.000Z' },
  { id: 'd-c0004', type: 'cotizacion', number: 4, date: '2026-10-19', clientId: 'c4', clientName: 'Mantenimiento Norte', attention: 'Ana Torres', email: 'ana@mantenimientonorte.mx', phone: '55 3011 7482', items: [{ id: 'i-c0004', quantity: 5, unit: 'pza', description: 'Silicón transparente', unitPrice: 115 }], delivery: 'Inmediata', payment: 'Por definir', note: '', createdAt: '2026-10-19T12:00:00.000Z' },
  { id: 'd-c0003', type: 'cotizacion', number: 3, date: '2026-10-18', clientId: 'c1', clientName: 'Constructora del Valle', attention: 'María López', email: 'compras@constructoravalle.com', phone: '55 1234 5678', items: [{ id: 'i-c0003', quantity: 8, unit: 'bulto', description: 'Cemento gris 50 kg', unitPrice: 285 }], delivery: 'Inmediata', payment: 'Por definir', note: '', createdAt: '2026-10-18T12:00:00.000Z' },
  { id: 'd-c0002', type: 'cotizacion', number: 2, date: '2026-10-17', clientId: 'c2', clientName: 'Retama', attention: 'Jorge Ramírez', email: 'jorge@retama.mx', phone: '55 4589 2210', items: [{ id: 'i-c0002', quantity: 10, unit: 'pza', description: 'Varilla corrugada 3/8', unitPrice: 145 }], delivery: 'Inmediata', payment: 'Por definir', note: '', createdAt: '2026-10-17T12:00:00.000Z' },
];

function read<T>(key: string, fallback: T): T { const value = localStorage.getItem(key); return value ? JSON.parse(value) as T : fallback; }
function write<T>(key: string, value: T) { localStorage.setItem(key, JSON.stringify(value)); }
export const storage = {
  getClients: () => read(keys.clients, starterClients), saveClients: (value: Client[]) => write(keys.clients, value),
  getProducts: () => read(keys.products, starterProducts), saveProducts: (value: Product[]) => write(keys.products, value),
  getDocuments: () => read(keys.documents, starterDocuments), saveDocuments: (value: BusinessDocument[]) => write(keys.documents, value),
  getSettings: () => read<BusinessSettings>(keys.settings, { name: 'Ferretería García', logo: '', salesName: 'Geovany Lanzagorta', salesPhone: '81 26 16 21 88', salesEmail: 'ferreteriagarciamy@gmail.com' }),
  saveSettings: (value: BusinessSettings) => write(keys.settings, value),
};
