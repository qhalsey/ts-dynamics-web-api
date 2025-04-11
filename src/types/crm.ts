// /types/crm.ts

export interface EntityView {
  name: string;
  description?: string;
  returnedtypecode: string;
  isdefault: boolean;
  ismanaged: boolean;
  fetchxml: string;
  layoutxml: string;
  layoutjson: string;
  iscustomizable: boolean;
}

export interface BusinessRule {
  name: string;
  description?: string;
  primaryentity: string;
  xaml?: string;
  clientdata?: string;
  scope: number;
  ismanaged: boolean;
  iscustomizable: boolean;
  statecode: number;
  statuscode: number;
  type: number;
  category: number;
}

export interface Relationship {
  "Schema Name": string;
  "Security Types": string;
  Managed: string;
  Type: string;
  "Attribute Ref.": string;
  "Entity Ref.": string;
  "Referencing Attribute": string;
  "Referencing Entity": string;
  Hierarchical: string;
  Behavior: string;
  Customizable: string;
  "Menu Behavior": string;
  "Menu Customization": string;
  Assign: string;
  Delete: string;
  Archive: string;
  Merge: string;
  Reparent: string;
  Share: string;
  Unshare: string;
  RollupView: string;
}
