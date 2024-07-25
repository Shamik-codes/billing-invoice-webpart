// File: src/webparts/billingInvoice/components/IBillingInvoiceProps.ts

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBillingInvoiceProps {
  description: string;
  context: WebPartContext;
}