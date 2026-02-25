<div align="center">
  <h1>🏢 SPFx Employees List</h1>
  <p>Modern, responsive, and dynamic Microsoft 365 / SharePoint user directory powered by SPFx & React 18.</p>

  <img src="https://img.shields.io/badge/SharePoint-0078D4?style=for-the-badge&logo=microsoftsharepoint&logoColor=white" alt="SharePoint"/>
  <img src="https://img.shields.io/badge/React_18-20232A?style=for-the-badge&logo=react&logoColor=61DAFB" alt="React 18"/>
  <img src="https://img.shields.io/badge/TypeScript-007ACC?style=for-the-badge&logo=typescript&logoColor=white" alt="TypeScript"/>
  <img src="https://img.shields.io/badge/Radix_UI-161618?style=for-the-badge&logo=radixui&logoColor=white" alt="Radix UI"/>
  <img src="https://img.shields.io/badge/@mustafaaksoy41/sharepoint--kit-FF6C37?style=for-the-badge" alt="SharePoint Kit"/>
</div>

<br/>

## 🎯 Overview

**SPFx Employees List** is a custom SharePoint Framework (SPFx) Web Part designed to fetch and display organizational members in a clean, modern card grid layout. It pulls data directly from SharePoint lists and enriches it dynamically by fetching user profile information via the **Microsoft Graph API**.

### ✨ Key Features
- **Smart Search & Filtering:** Instantly search through team members or filter them directly by their `Job Title`.
- **Modern UI & Aesthetics:** Built completely with [Radix UI](https://www.radix-ui.com/) primitives and themes for an accessible, beautiful, and consistent interface.
- **Dynamic Profile Modals:** Clicking on a user reveals a detailed profile card fetching extended Azure AD information (Department, Phone, Location) using Graph API.
- **Scroll Optimization:** Custom layout prevents the web part container from endlessly stretching 100vh vertically, utilizing a max-height strategy with internal scrolling.

---

## ⚛️ React 18 Compatibility

Although Microsoft's default SPFx generators still often tie you down to React 16 or 17 setups, this project has been carefully configured to be fully compatible with the modern **React 18** ecosystem. 

This enables the use of concurrent rendering practices, modern hook optimizations, and ensures compatibility with the latest generation of React UI libraries (like Radix UI) without throwing legacy version conflicts.

---

## 🛠 Powered by `@mustafaaksoy41/sharepoint-kit`

This project heavily leverages the open-source **[sharepoint-kit](https://www.npmjs.com/package/@mustafaaksoy41/sharepoint-kit)** library to dramatically reduce boilerplate code and handle SPFx complexities elegantly:

1. **Context & Auth (`<SpProvider>`)**: Automatically manages MS Graph API and SPHttpClient tokens, propagating them down the React tree without prop-drilling.
2. **Resilience (`<SpErrorBoundary>`)**: Catches internal SharePoint API / token errors gracefully and prevents the whole web part from crashing.
3. **Type-Safety (`sp-generate-types`)**: By utilizing the built-in CLI, the project reads `config/sharepoint.config.ts`, connects to the actual SharePoint tenant, and auto-generates TypeScript models (like `Employee` or `Invoice`). This guarantees that your React code perfectly matches your live SharePoint List columns!

---

## 🚀 Getting Started

Follow these steps to spin up the project locally on your SharePoint tenant.

### 1. Environment Variables Setup
The project protects sensitive IDs through local environment files.

1. Copy the example file:
   ```bash
   cp .env.example .env
   ```
2. Open `.env` and fill in your actual Microsoft 365 / Azure AD details:
   ```env
   SHAREPOINT_SITE_ID=root
   SHAREPOINT_TENANT_ID=xxxx-xxxx-xxxx
   SHAREPOINT_CLIENT_ID=yyyy-yyyy-yyyy
   SHAREPOINT_CLIENT_SECRET=zzzzzzzzzzz
   SHAREPOINT_INVOICE_LIST_ID=your-invoice-list-guid
   SHAREPOINT_EMPLOYEE_LIST_ID=your-employee-list-guid
   ```

### 2. SharePoint Configuration
1. Copy the example config file:
   ```bash
   cp config/sharepoint.config.example.ts config/sharepoint.config.ts
   ```
2. The config file will automatically read the IDs from your `.env` file, but you can further customize your list structures and output type names inside this TypeScript config.

### 3. Generate Models (Optional but Recommended)
Sync your local TypeScript interfaces with your live SharePoint lists:
```bash
npx sp-generate-types --config config/sharepoint.config.ts
```

### 4. Build and Serve
```bash
npm install
npm run serve
```
> The local workbench will boot up. You can add the web part to the page and test it!

---

## 📦 Roadmap: NPM Package Integration

*What's next?*
Currently, this is maintained as a standalone SPFx project. In the near future, the core web part logic and UI components will be extracted, generalized, and published as a reusable `npm` package. 

This will allow developers to simply `npm install @mustafaaksoy41/spfx-employees-list` inside any new SPFx project and mount the component with just a few props!

---

<div align="center">
  <i>Developed with ❤️ by <a href="https://github.com/mustafaaksoy41">Mustafa Aksoy</a></i>
</div>