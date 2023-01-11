import * as React from 'react';
import styles from './WebPartTabs.module.scss';
import { IWebPartTabsProps } from './IWebPartTabsProps';
import { Pivot, PivotItem, Label } from '@fluentui/react'

export default function WebPartTabs({ tabStyle, wpContext, collectionTabs, displayMode }: IWebPartTabsProps) {
  const [selectedKey, setSelectedKey] = React.useState(() => {
    if (collectionTabs.length === 0) return null
    return collectionTabs[0].WebPart
  })

  const handleHideWebPart = (webPartIdShown: string) => {
    const allWebPartOnPage = document.querySelectorAll('div[class*="ControlZone ControlZone--clean"]')
    allWebPartOnPage.forEach(item => {
      const webPartContent = item.querySelector('div[data-sp-web-part-id*="-"]')
      if (
        webPartContent.getAttribute('data-sp-web-part-id') === webPartIdShown ||
        webPartContent.getAttribute('data-sp-web-part-id') === wpContext.manifest.id
      ) {
        if (!item.classList.contains(styles.show)) item.classList.add(styles.show)
        if (item.classList.contains(styles.hide)) item.classList.remove(styles.hide)
      } else {
        if (!item.classList.contains(styles.hide)) item.classList.add(styles.hide)
        if (item.classList.contains(styles.show)) item.classList.remove(styles.show)
      }
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
    handleHideWebPart(selectedKey)
  }, [selectedKey, displayMode])

  const handleLinkClick = (item?: PivotItem) => {
    if (item) {
      setSelectedKey(item.props.itemKey!);
    }
  }

  return (
    <>
      <Pivot
        aria-label="Custom Pivot"
        linkFormat={tabStyle}
        headersOnly={true}
        selectedKey={selectedKey}
        onLinkClick={handleLinkClick}
      >
        {
          collectionTabs.map(item => (
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
