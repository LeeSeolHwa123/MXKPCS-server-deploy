import assert from "node:assert/strict";
import { findDuplicateRegistration } from "../api/xrf-sharepoint.js";

const itemRows=[
  {id:"item-1",fields:{Title:"ITM_A_100",field_1:"A-100"}},
];
const requestRows=[
  {id:"request-1",fields:{Title:"REQ_NEW_20261016_001",field_1:"New",field_2:"Pending",field_4:"ITM_B_200",field_5:"B-200"}},
  {id:"request-2",fields:{Title:"REQ_NEW_20261016_002",field_1:"New",field_2:"Rejected",field_4:"ITM_C_300",field_5:"C-300"}},
];

assert.equal(findDuplicateRegistration(itemRows,requestRows," a-100 ","ITM_A_100")?.source,"XRF_Items");
assert.equal(findDuplicateRegistration(itemRows,requestRows,"b-200","ITM_B_200")?.source,"XRF_Requests");
assert.equal(findDuplicateRegistration(itemRows,requestRows,"C-300","ITM_C_300"),null);
assert.equal(findDuplicateRegistration(itemRows,requestRows,"D-400","ITM_D_400"),null);

console.log("registration-duplicate: PASS");
