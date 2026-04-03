// subscribedSkus.types.ts

export interface SubscribedSkusResponse {
  '@odata.context'?: string;
  value: SubscribedSku[];
}

export interface SubscribedSku {
  accountName?: string;
  accountId?: string;
  appliesTo: string;
  capabilityStatus: string;
  consumedUnits: number;
  id: string;
  prepaidUnits: LicenseUnitsDetail;
  servicePlans: ServicePlanInfo[];
  skuId: string;
  skuPartNumber: string;
  subscriptionIds?: string[];
}

export interface LicenseUnitsDetail {
  enabled: number;
  suspended: number;
  warning: number;
  lockedOut: number;
}

export interface ServicePlanInfo {
  servicePlanId: string;
  servicePlanName: string;
  provisioningStatus: string;
  appliesTo: string;
}
