import { ComboBox, IComboBoxProps } from '@fluentui/react/lib/ComboBox';
import * as React from 'react';
import { FC } from 'react';
import { Controller } from 'react-hook-form';
import { HookFormProps } from '../HookFormProps';

export const ControlledComboBox: FC<HookFormProps & IComboBoxProps> = (props) => {
  return (
    <Controller
      name={props.name}
      control={props.control}
      rules={props.rules}
      defaultValue={props.defaultValue || ''}
      render={({ onChange, onBlur, value, name: fieldName }) => {
        console.log(value);
        return (
          <ComboBox
            {...props}
            selectedKey={value}
            onChange={(_, option) => {
              onChange(option.key);
            }}
            onBlur={onBlur}
            errorMessage={props.errors[fieldName] && props.errors[fieldName].message}
            defaultValue={undefined}
            multiSelect
        />
        )}}
    />
  );
};