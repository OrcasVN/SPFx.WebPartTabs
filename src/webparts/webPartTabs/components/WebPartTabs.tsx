import * as React from 'react';
import styles from './WebPartTabs.module.scss';
import { IWebPartTabsProps } from './IWebPartTabsProps';
import { Pivot, PivotItem, Label } from '@fluentui/react'

export default function WebPartTabs({ }: IWebPartTabsProps) {
  return (
    <>
      <Pivot aria-label="Basic Pivot Example">
        <PivotItem
          headerText="My Files"
          headerButtonProps={{
            'data-order': 1,
            'data-title': 'My Files Title',
          }}
        >
          <Label >Pivot #1</Label>
        </PivotItem>
        <PivotItem headerText="Recent">
          <Label>Pivot #2</Label>
        </PivotItem>
        <PivotItem headerText="Shared with me">
          <Label>Pivot #3</Label>
        </PivotItem>
      </Pivot>
    </>
  )
}
