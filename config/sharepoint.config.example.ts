import type { SpKitConfig } from '@mustafaaksoy41/sharepoint-kit';

const config: SpKitConfig = {
  // --- TEMEL BİLGİLER & BAĞLANTI ---
  // Uygulamanın çalışması için gereken temel ortam değişkenleri
  siteId: process.env.SHAREPOINT_SITE_ID || 'root',
  tenantId: process.env.SHAREPOINT_TENANT_ID,
  clientId: process.env.SHAREPOINT_CLIENT_ID,
  clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
  defaultStrategy: 'interactive',

  // --- 1. BELGELER ve ŞABLONLAR (Content Types) ---
  // Doküman kütüphaneleri veya özel form şablonları için kullanılır
  contentTypes: [
    {
      listId: process.env.SHAREPOINT_INVOICE_LIST_ID || '', // Örnek Belge
      contentTypeName: 'Fatura Denemesi',
      outputType: 'Invoice',
    }
  ],

  // --- 2. DOĞRUDAN LİSTELER (Lists) ---
  // İçerisinde herhangi bir özel Content Type barındırmayan, düz tablolar için kullanılır
  lists: [
    {
      listId: process.env.SHAREPOINT_EMPLOYEE_LIST_ID || '', // Doğrudan Listenin ID'si
      outputType: 'Employee',                          // Oluşacak TS arayüzü (örn: export interface Employee)
    }
  ],

  // --- ÇIKTI VE HARİTALAMA AYARLARI ---
  options: {
    outputDir: './src/models',
    
    // SharePoint'in arka plandaki bozuk veya boşluklu Türkçe kolon isimlerini
    // kod tarafında düzgün (camelCase) okumak için kullanılır
    fieldNameMapping: {
      'JobTitle': 'jobTitle',
      'Title': 'title',
      'PersonLookupId': 'userId'
    },
  },
};

export default config;
