import { SPFI, SPFx} from "@pnp/sp";

import { getSP } from "../pnpjsconfig";
import { IMolaceContractRequestFormLmSolutionProps } from "../webparts/molaceContractRequestFormLmSolution/components/IMolaceContractRequestFormLmSolutionProps";
import * as formconst from "../webparts/constant";
import { Web } from "@pnp/sp/webs";



  export const updateData=(props:IMolaceContractRequestFormLmSolutionProps ,itemId: any, data: any): Promise<void>=> {
    const _sp :SPFI = getSP(props.context) ;
    return new Promise<void>((resolve, reject) => {
      _sp.web.lists.getByTitle(formconst.LISTNAME).items.getById(itemId).update(data)
        .then(() => {
          
          //console.log(e.response.headers.get("content-length"))
          resolve();
        })
        .catch((error) => {
   
          reject(error);
        });
    });
  }

  export const getCustomerTitle=(props:IMolaceContractRequestFormLmSolutionProps,customerRef: string) => {
    //console.log(customerName)
    const _web = Web(formconst.CUSTOMER_URL).using(SPFx(props.context));
    return new Promise((resolve, reject) => {
      _web.lists.getByTitle(formconst.CUSTOMER_LISTNAME).items.select("Title").filter(`RefCode eq '${customerRef}'`)()
        .then((items) => {
          if (items.length > 0) {
            const customerTitle = items[0].Title;
            
            resolve(customerTitle);
          } else {
            reject(new Error("Customer not found"));
          }
        })
        .catch((error) => {
          reject(error);
        });
    });
  } 
  

