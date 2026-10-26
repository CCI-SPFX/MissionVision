import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; 
import { IMissionVisionProps, IGridItem } from './IMissionVisionProps';
import "./style.css";
const MissionVision: React.FunctionComponent<IMissionVisionProps> = (props) => {
  const {
    listName,
    siteUrl,
    spHttpClient, 
    alignment,
    gridBackground,
    itemWidth,
    gridGap,
    itemAsLink,
    itemBackground,
    iconWidth,
    titleFontSize,
    descFontSize
  } = props;

  const [items, setItems] = useState<IGridItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');

  useEffect(() => {
    let isMounted = true;

    const loadItems = async (): Promise<void> => {
      if (!listName) {
        if (isMounted) {
          setLoading(false);
          setError('Please enter a list name in the web part properties.');
        }
        return;
      }

      setLoading(true);
      setError('');

      try { 
        const itemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$select=Id,Title,Description,URL`;
        const itemsResponse: SPHttpClientResponse = await spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1);

        if (!itemsResponse.ok) {
          throw new Error(`Failed to load list items (${itemsResponse.status} ${itemsResponse.statusText})`);
        }

        const itemsJson = await itemsResponse.json();
        const rawItems: any[] = itemsJson.value || [];
 
        const gridItems: IGridItem[] = await Promise.all(
          rawItems.map(async (item) => {
            let iconUrl = '';
            try {
              const attachUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${item.Id})/AttachmentFiles`;
              const attachResponse: SPHttpClientResponse = await spHttpClient.get(attachUrl, SPHttpClient.configurations.v1);

              if (attachResponse.ok) {
                const attachJson = await attachResponse.json();
                const attachments: any[] = attachJson.value || [];
                if (attachments.length > 0) {
                  iconUrl = attachments[0].ServerRelativeUrl;
                }
              }
            } catch {
              // attachment fetch errors for this item, show no icon
            }
 
            const linkUrl = item.URL && typeof item.URL === 'object' ? item.URL.Url : (item.URL || '#');

            return {
              id: item.Id,
              title: item.Title || '',
              description: item.Description || '',
              url: linkUrl,
              iconUrl
            } as IGridItem;
          })
        );

        if (isMounted) {
          setItems(gridItems);
          setLoading(false);
        }
      } catch (err: any) {
        if (isMounted) {
          setError(err.message || 'Something went wrong while loading the list.');
          setLoading(false);
        }
      }
    };

    loadItems();

    return () => {
      isMounted = false;
    };
  }, [listName, siteUrl]);

  if (loading) {
    return <div>Loading...</div>;
  }

  if (error) {
    return <div className="error">{error}</div>;
  }

  if (items.length === 0) {
    return <div>No items found in this list.</div>;
  }

const gridContainerStyle: React.CSSProperties = {
  backgroundColor: gridBackground || '#8BC9BA',
  gap: gridGap || '1px',
  gridTemplateColumns: `repeat(auto-fit, minmax(${itemWidth || '195px'}, 1fr))`
};

  const alignItem: Record<string, React.CSSProperties['textAlign']> = { left: 'left', center: 'center', right: 'right' };
const alignMap: Record<string, string> = { left: 'flex-start', center: 'center', right: 'flex-end' };
  const itemStyle: React.CSSProperties = {
    backgroundColor: itemBackground || "#323E48",
    alignItems: alignMap[alignment] || 'center',
    textAlign: alignItem[alignment] || 'center'
  }; 

  const iconStyle: React.CSSProperties = {
    width: iconWidth || "80px"
  };

  const titleStyle: React.CSSProperties = {
    fontSize: titleFontSize || "17px"
  };

  const descStyle: React.CSSProperties = {
    fontSize: descFontSize || "14px"
  };

  return (
    <div className="gridContainer" style={gridContainerStyle}>
      {items.map((item) => {
        const content = (
          <React.Fragment>
            <div className="icon">
              {item.iconUrl && (
                <img src={item.iconUrl} alt="Icon" style={iconStyle} />
              )}
            </div>
            <h3 style={titleStyle}>{item.title}</h3>
            <p style={descStyle}>{item.description}</p>
          </React.Fragment>
        );

        return itemAsLink ? (
          <a href={item.url} className="gridItem" key={item.id} style={itemStyle}>
            {content}
          </a>
        ) : (
          <div className="gridItem" key={item.id} style={itemStyle}>
            {content}
          </div>
        );
      })}
    </div>
  );
};

export default MissionVision;