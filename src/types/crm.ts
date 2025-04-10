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
