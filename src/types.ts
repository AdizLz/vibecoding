export type DocumentType = 'remision' | 'cotizacion';
export type Unit = 'pza' | 'kg' | 'lt' | 'm' | 'caja' | 'bulto' | 'juego';

export interface Client {
  id: string;
  name: string;
  attention: string;
  email: string;
  phone: string;
}

export interface Product {
  id: string;
  name: string;
  unit: Unit;
  price: number;
}

export interface LineItem {
  id: string;
  quantity: number;
  unit: Unit;
  description: string;
  unitPrice: number;
}

export interface BusinessSettings {
  name: string;
  logo: string;
  salesName: string;
  salesPhone: string;
  salesEmail: string;
}

export interface BusinessDocument {
  id: string;
  type: DocumentType;
  number: number;
  date: string;
  clientId: string;
  clientName: string;
  attention: string;
  email: string;
  phone: string;
  items: LineItem[];
  delivery: string;
  payment: string;
  note: string;
  createdAt: string;
}

export const units: Unit[] = ['pza', 'kg', 'lt', 'm', 'caja', 'bulto', 'juego'];
export const money = (value: number) => value.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });
export const documentTitle = (type: DocumentType) => type === 'remision' ? 'Remisión' : 'Cotización';
export const documentFolio = (document: Pick<BusinessDocument, 'type' | 'number'>) => `${document.type === 'remision' ? 'R' : 'C'}${document.type === 'remision' ? document.number : String(document.number).padStart(4, '0')}`;
export const cleanDescription = (value: string) => {
  const cleaned = value.replace(/\.+/g, ' ').replace(/\s+/g, ' ').replace(/(\d)(lt|kg|pza|caja|bulto|juego|m)\b/gi, '$1 $2').trim().toLowerCase();
  return cleaned ? cleaned.charAt(0).toUpperCase() + cleaned.slice(1) : '';
};
export const subtotalOf = (document: Pick<BusinessDocument, 'items'>) => document.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
