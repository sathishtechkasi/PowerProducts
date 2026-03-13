 
import * as React from 'react';
import { Checkbox, Dropdown, IDropdownOption, Stack } from '@fluentui/react';
export interface IFieldOrderRendererProps {
  fieldKey: string;
  fieldTitle: string;
  selected: boolean;
  order: number;
  onToggle: (key: string, checked: boolean) => void;
  onOrderChange: (key: string, newOrder: number) => void;
  max: number;
}

export const FieldOrderRenderer: React.FC<IFieldOrderRendererProps> = (props) => {
  const options: IDropdownOption[] = Array.from({ length: props.max }, (_, i) => ({
    key: i + 1,
    text: (i + 1).toString()
  }));

  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { marginBottom: 8 } }}>
      <Checkbox
        label={props.fieldTitle}
        checked={props.selected}
        onChange={(_, checked) => props.onToggle(props.fieldKey, !!checked)}
      />
      <Dropdown
        selectedKey={props.order}
        options={options}
        onChange={(_, opt) => props.onOrderChange(props.fieldKey, opt?.key as number)}
        styles={{ root: { width: 60 } }}
      />
    </Stack>
  );
};