/**
 * 
 * Official Community Imports
 * 
 */

import * as React from 'react';

import { Slider, ISliderProps } from 'office-ui-fabric-react/lib/Slider';

export function createSlider( label: string , timeSliderValue , timeSliderMin, timeSliderMax, timeSliderInc, _onChange, disabled: boolean, minWidth: number = 350, hStyles : any = null ){

  return (
    <div style={{minWidth: minWidth }}>
      <Slider 
  //      label={ ((timeSliderValue < 0)  ? "Start time is in the past" : "End time is Back to the future" ) }  //This is the label to left of slider
        label = { label }
        min={ timeSliderMin } 
        max={ timeSliderMax } 
        step={ timeSliderInc } 
        defaultValue={ timeSliderValue } 
        disabled={ disabled }
        valueFormat={ value => timeSliderValue }  //This is the label on right of slider showing current value
  //      valueFormat = { null }
        showValue 
        originFromZero
        onChange={_onChange}
        styles = { hStyles }
     />

    </div>

  );

}

export function createChoiceSlider( label: string , timeSliderValue , timeSliderMax, timeSliderInc, _onChange, hStyles : any = null ){

  return (
    <div style={{minWidth: 250 }}>
      <Slider 
  //      label={ ((timeSliderValue < 0)  ? "Start time is in the past" : "End time is Back to the future" ) }  //This is the label to left of slider
        label = { label }
        min={ 0 } 
        max={ timeSliderMax } 
        step={ timeSliderInc } 
        defaultValue={ 0 } 
        valueFormat={ value => `${timeSliderValue}` }  //This is the label on right of slider showing current value
  //      valueFormat = { value => timeSliderValue }
        showValue 
        originFromZero
        onChange={_onChange}
        styles = { hStyles }
     />

    </div>

  );

}

export const verticalStyles = {
  container: { height: 100},//color: 'green' works here
};

export function createVerticalSlider( timeSliderValue , timeSliderMax, timeSliderInc, _onChange, vStyles : any = null ){

  /*
  let vStyles = {
    container: { height: height},//color: 'green' works here
  };
*/

  return (
    <div style={{minWidth: 250 }}>
      <Slider 
  //      label={ ((timeSliderValue < 0)  ? "Start time is in the past" : "End time is Back to the future" ) }  //This is the label to left of slider
        label = { 'Slide to adjust date range' }
        min={ 0 } 
        max={ timeSliderMax } 
        step={ timeSliderInc } 
        defaultValue={ 0 } 
        valueFormat={ value => `Offset ${value} px?`}  //This is the label on right of slider showing current value
  //      valueFormat = { null }
        showValue 
        originFromZero
        onChange={_onChange}
        vertical= { true }
        styles={vStyles}
     />
    </div>

  );

}



/*
function _onChange(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
  console.dir(option);
}
*/