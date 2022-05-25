import { DatePicker, defaultDatePickerStrings, IDatePickerProps } from '@fluentui/react/lib/DatePicker';
import * as React from 'react';
import { FC } from 'react';
import { Controller } from 'react-hook-form';
import { HookFormProps } from '../HookFormProps';

export const ControlledDatePicker: FC<HookFormProps & IDatePickerProps> = (props) => {


  const onFormatDate = (date?: Date): string => {
    return !date ? '' : (('0' + (date.getMonth() + 1)).slice(-2) + '/' + ('0' + (date.getDate())).slice(-2) + '/' + date.getFullYear());
  };



  return (
    <Controller
      name={props.name}
      control={props.control}
      rules={props.rules}
      defaultValue={props.defaultValue || ''}
      render={({ onChange, onBlur, value, name: fieldName }) => (
        <DatePicker
          {...props}
          value={value}
          onSelectDate={(date: Date) => {
            if (!date) {
               onChange('');
            } else {
              let dateVal = ('0' + (date.getMonth() + 1)).slice(-2) + '/' + ('0' + (date.getDate())).slice(-2) + '/' + date.getFullYear();
               onChange(dateVal);
            }
          }}
          // onChange={value}
          onBlur={onBlur}
          formatDate={onFormatDate}
          // parseDateFromString={onParseDateFromString}
          // className={styles.control}
          // DatePicker uses English strings by default. For localized apps, you must override this prop.
          strings={defaultDatePickerStrings}
          defaultValue={undefined}
        />
      )}
    />
  );
};