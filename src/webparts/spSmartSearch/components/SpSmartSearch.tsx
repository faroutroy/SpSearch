const ResultItemRow: React.FC<{ item: ISpSmartItem }> = ({ item }) => (
  <li className={styles.resultItem}>
    <div className={styles.resultIcon}>
      {item.type === 'listItem' ? '📋' : getFileIcon(item.fileType)}
    </div>
    <div className={styles.resultContent}>
      <a href={item.url} target="_blank" rel="noopener noreferrer"
        className={styles.resultTitle} title={item.title}>
        {item.title}
      </a>
      <div className={styles.resultMeta}>
        {item.fileType && <span>📄 {item.fileType.toUpperCase()}</span>}
        {item.listName && <span>📂 {item.listName}</span>}
        {item.author && <span>👤 {item.author}</span>}
        {item.modifiedDate && <span>📅 Modified {formatDate(item.modifiedDate)}</span>}
      </div>
      <div className={styles.resultMeta}>
        {item.businessArea && <span>🏢 {item.businessArea}</span>}
        {item.segment && <span>🔖 {item.segment}</span>}
        {item.project && <span>📋 {item.project}</span>}
        {item.estimator && <span>👷 {item.estimator}</span>}
        {item.bid2WinId && <span>🆔 {item.bid2WinId}</span>}
        {item.bidDate && <span>📆 Bid: {formatDate(item.bidDate)}</span>}
      </div>
      <div className={styles.resultMeta}>
        {item.city && <span>🏙 {item.city}</span>}
        {item.county && <span>🗺 {item.county}</span>}
        {item.state && <span>📍 {item.state}</span>}
        {item.zipCode && <span>📮 {item.zipCode}</span>}
        {item.owner && <span>🏠 {item.owner}</span>}
      </div>
      <div className={styles.resultMeta}>
        {item.sqYards && <span>📐 {item.sqYards} sq yds</span>}
        {item.laneMiles && <span>🛣 {item.laneMiles} lane mi</span>}
        {item.numberOfLots && <span>🔢 {item.numberOfLots} lots</span>}
      </div>
      {item.summary && (
        <div className={styles.resultSummary}
          dangerouslySetInnerHTML={{ __html: item.summary }} />
      )}
    </div>
  </li>
);
