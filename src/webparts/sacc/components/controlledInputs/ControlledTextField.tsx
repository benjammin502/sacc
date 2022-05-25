import { ITextFieldProps, TextField } from '@fluentui/react/lib/TextField';
import * as React from 'react';
import { FC } from 'react';
import { Controller } from 'react-hook-form';
import { HookFormProps } from '../HookFormProps';

export const ControlledTextField: FC<HookFormProps & ITextFieldProps> = (props) => {
  return (
    <Controller
      name={props.name}
      control={props.control}
      rules={props.rules}
      defaultValue={props.defaultValue || ''}
      render={({ onChange, onBlur, value, name: fieldName }) => (
        <TextField
          {...props}
          onChange={onChange}
          value={value}
          onBlur={onBlur}
          name={fieldName}
          errorMessage={props.errors[fieldName] && props.errors[fieldName].message}
          defaultValue={undefined}
        />
      )}
    />
  );
};