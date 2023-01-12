import * as React from 'react';
import styles from './WebPartTabs.module.scss';
import { IWebPartTabsProps } from './IWebPartTabsProps';
import { Pivot, PivotItem, Label } from '@fluentui/react'

export default function WebPartTabs({ tabStyle, wpContext, collectionTabs, displayMode, fontSize }: IWebPartTabsProps) {
  const [selectedKey, setSelectedKey] = React.useState(() => {
    if (collectionTabs.length === 0) return null
    return collectionTabs[0].WebPart
  })

  const handleHideWebPart = (webPartIdShown: string) => {
    const allWebPartOnPage = document.querySelectorAll('div[class*="ControlZone ControlZone--clean"]')
    allWebPartOnPage.forEach(item => {
      if (
        item.id === webPartIdShown ||
        item.id === wpContext.instanceId
      ) {
        if (!item.classList.contains(styles.show)) item.classList.add(styles.show)
        if (item.classList.contains(styles.hide)) item.classList.remove(styles.hide)
      } else {
        if (!item.classList.contains(styles.hide)) item.classList.add(styles.hide)
        if (item.classList.contains(styles.show)) item.classList.remove(styles.show)
      }
    })
  }

  const showAllWebPart = () => {
    const allWebPartOnPage = document.querySelectorAll('div[class*="ControlZone ControlZone--clean"]')
    allWebPartOnPage.forEach(item => {
      if (item.classList.contains(styles.show)) item.classList.remove(styles.show)
      if (item.classList.contains(styles.hide)) item.classList.remove(styles.hide)
    })
  }

  React.useEffect(() => {
    (async () => {
      let newSelectedKey = selectedKey
      if (collectionTabs.filter(item => item.WebPart === selectedKey).length <= 0) newSelectedKey = collectionTabs[0]?.WebPart || null

      setSelectedKey(newSelectedKey)
    })().then(() => { }).catch(e => console.log(e))
  }, [wpContext, collectionTabs])

  React.useEffect(() => {
    if (displayMode === 1) handleHideWebPart(selectedKey)
  }, [selectedKey])

  React.useEffect(() => {
    if (displayMode === 1) {
      handleHideWebPart(selectedKey)
    } else {
      showAllWebPart()
    }
  }, [displayMode])

  const handleLinkClick = (item?: PivotItem) => {
    if (item) {
      setSelectedKey(item.props.itemKey!);
    }
  }

  const handleFontSize = (fontSizeString: string): string => {
    const removedSpacing = fontSizeString ? fontSizeString.replace(/[ ]/g, '') : '14px'
    return isNaN(Number(removedSpacing)) ? removedSpacing : `${Number(removedSpacing)}px`
  }

  return (
    <>
      <Pivot
        aria-label="Custom Pivot"
        linkFormat={tabStyle}
        headersOnly={true}
        selectedKey={selectedKey}
        onLinkClick={handleLinkClick}
        styles={{
          text: { fontSize: handleFontSize(fontSize) },
          icon: { fontSize: handleFontSize(fontSize) }
        }}
      >
        {
          collectionTabs.sort((a, b) => Number(a.DisplayOrder) - Number(b.DisplayOrder)).map(item => (
            <PivotItem
              headerText={item.Title}
              itemKey={item.WebPart}
              itemIcon={item.IconName || 'none'}
              key={item.uniqueId}
            >
            </PivotItem>
          ))
        }
      </Pivot>
    </>
  )
}
