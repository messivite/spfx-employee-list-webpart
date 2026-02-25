import type { SpKitConfig } from '@mustafaaksoy41/sharepoint-kit';

const config: SpKitConfig = {
  // Genel bağlantı ayarları
  // Uygulamanın çalışması için gereken temel ortam değişkenleri
  siteId: process.env.SHAREPOINT_SITE_ID || 'root',
  tenantId: process.env.SHAREPOINT_TENANT_ID,
  clientId: process.env.SHAREPOINT_CLIENT_ID,
  clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
  defaultStrategy: 'interactive',

  // Doküman kütüphaneleri - Content Types
  contentTypes: [
    {
      listId: process.env.SHAREPOINT_INVOICE_LIST_ID || '', // Örnek Belge
      contentTypeName: 'Fatura Denemesi',
      outputType: 'Invoice',
    }
  ],

  // Doğrudan Liste tanımlamaları - Content types barındırmayan normal listeler
  lists: [
    {
      listId: process.env.SHAREPOINT_EMPLOYEE_LIST_ID || '', // Doğrudan Listenin ID'si
      outputType: 'Employee',                          // Oluşacak TS arayüzü (örn: export interface Employee)
    }
  ],

  // Çıktı ve Field mapping ayarları
  options: {
    outputDir: './src/models',
    
    // Type hatalarını önlemek için SharePoint isimlerindeki boşluk/Türkçe karakter vs. haritalaması
    fieldNameMapping: {
      'JobTitle': 'jobTitle',
      'Title': 'title',
      'PersonLookupId': 'userId'
    },
  },
};

export default config;
