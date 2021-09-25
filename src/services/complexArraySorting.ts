
import { ISeriesSort } from '@mikezimm/npmfunctions/dist/CSSCharts/ICSSCharts';

//2021-01-05: Updated per TrackMyTime7 arrayServices
export function sortObjectArrayByNumberKey( arr: any[], order: ISeriesSort, key: string ) : any[] {

  let keys = key.split('.');
  let key1 = keys.length >= 1 ? keys[0] : key;
  let key2 = keys.length >= 2 ? keys[1] : '';
  let key3 = keys.length >= 3 ? keys[2] : '';

  if ( keys.length === 1 ) {
    if ( order === 'asc' ) { 
      arr.sort((a, b) => a[key]-b[key] );
    } else if ( order === 'dec' ) {
        arr.sort((a, b) => b[key]-a[key] );
    } else {
        
    }
  } else if ( keys.length === 2 ) {
    if ( order === 'asc' ) { 
      arr.sort((a, b) => a[key1][key2]-b[key1][key2] );
    } else if ( order === 'dec' ) {
        arr.sort((a, b) => b[key1][key2]-a[key1][key2] );
    } else {
        
    }
  }

  return arr;

}