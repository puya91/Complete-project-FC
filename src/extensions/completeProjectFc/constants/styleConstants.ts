import { IDropdownStyles, IStackTokens, mergeStyleSets } from "@fluentui/react";

export const dropdown: Partial<IDropdownStyles> = { dropdown: { maxWidth: 300, minWidth: 200, width: 300 } };
export const stackTokens: IStackTokens = { childrenGap: 30, padding: 15 };
export const buttonStackTokens: IStackTokens = { childrenGap: 15, padding: 15 };
export const verticalGapStackTokens: IStackTokens = { padding: 20 };
export const centerVerticalAlignItems = {
  root: [{
      alignItems: 'center',
  }]
};
export const inputStyle = {
  root: [{
      maxWidth: '300px',
      width: '300px',
      minWidth: '200px',
  }]
};
export const inputStyleLarge = {
  root: [{
      maxWidth: '630px',
      minWidth: '200px',
  }]
};
export const filePickerButtonUploadStyle = {
  root: [{
      color: 'black',
      fontWeight: '800',
      fontSize: '18px',
  }]
}
export const popupStyles = mergeStyleSets({
  root: {
      background: 'rgba(0, 52, 120, .85)',
      bottom: '0',
      left: '0',
      position: 'fixed',
      right: '0',
      top: '0',
      boxShadow: '0px 8px 20px rgba(0, 52, 120, 0.15)',
  },
  content: {
      background: 'white',
      left: '50%',
      maxWidth: '464px',
      width: '464px',
      height: '212px',
      padding: '24px 48px 32px 48px',
      position: 'absolute',
      top: '50%',
      transform: 'translate(-50%, -50%)',
      borderRadius: '8px',
      fontSize: '16px',
      lineHeight: '24px',
      weight: '400',
  },
})
export const confirmTitleStyle = {
  display: 'inline-block',
  verticalAlign: 'middle',
  fontWeight: '500',
  fontSize: '25px',
  lineHeight: '30px',
  color: 'black',
  width: '436px',
  margin: '0',
}