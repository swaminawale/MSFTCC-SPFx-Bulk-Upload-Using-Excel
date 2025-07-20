import * as React from "react";
import styles from "./BulkUploadSpFx.module.scss";
import type {
  IBulkUploadSpFxProps,
  ISharePointItem,
} from "./IBulkUploadSpFxProps";
import * as XLSX from "xlsx";
import {
  getItemsFromList,
  saveDataBatch,
  saveDataSequential,
} from "./Services/Services";
import {
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  Text,
  Stack,
} from "@fluentui/react";

export const BulkUploadSpFx: React.FC<IBulkUploadSpFxProps> = ({ context }) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [items, setItems] = React.useState<ISharePointItem[]>([]);
  const [processedItems, setProcessedItems] = React.useState<ISharePointItem[]>(
    []
  );
  const [missingRows, setMissingRows] = React.useState<
    { row: number; fields: string[] }[]
  >([]);
  const [loading, setLoading] = React.useState(true);
  const [uploading, setUploading] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const [sentCount, setSentCount] = React.useState(0);
  const [error, setError] = React.useState<string | null>(null);

  const fetchList = React.useCallback(async () => {
    setLoading(true);
    try {
      const listData = await getItemsFromList(context, "Bulk Upload List");
      setItems(listData);
    } catch {
      setError("Unable to load list items.");
    } finally {
      setLoading(false);
    }
  }, [context]);

  React.useEffect(() => {
    void fetchList();
  }, [fetchList]);

  const handleFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ({ target }) => {
      const data = target?.result;
      if (!data) return;
      const wb = XLSX.read(data, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw: any[] = XLSX.utils.sheet_to_json(ws, { raw: true });

      const misses: { row: number; fields: string[] }[] = [];
      const valid: ISharePointItem[] = [];

      raw.forEach((r, i) => {
        const missing: string[] = [];
        ["FirstName", "LastName", "WorkEmail", "HireDate", "Salary"].forEach(
          (key) => {
            if (!r[key]) missing.push(key);
          }
        );
        if (missing.length) misses.push({ row: i + 2, fields: missing });
        else
          valid.push({
            Title: r.Title || `${r.FirstName} ${r.LastName}`,
            FirstName: r.FirstName,
            LastName: r.LastName,
            WorkEmail: r.WorkEmail,
            PersonalEmail: r.PersonalEmail,
            BirthDate: r.BirthDate,
            HireDate: r.HireDate,
            WorkMode: r.WorkMode,
            Salary: r.Salary,
            IsMarried: r.IsMarried === "Yes",
            SocialProfile: {
              Url: r.SocialProfile?.Url || "",
            },
            JobTitle: r.JobTitle,
            About: r.About,
          });
      });

      setMissingRows(misses);
      setProcessedItems(misses.length ? [] : valid);
      setSentCount(0);
      setProgress(0);
      if (fileInputRef.current) fileInputRef.current.value = "";
    };
    reader.readAsArrayBuffer(file);
  };

  const runUpload = async (useBatch: boolean) => {
    if (!processedItems.length) return;
    setUploading(true);
    setProgress(0);
    setSentCount(0);
    setError(null);
    const fn = useBatch ? saveDataBatch : saveDataSequential;
    try {
      await fn(processedItems, context, "Bulk Upload List", (sent, total) => {
        setSentCount(sent);
        setProgress(sent / total);
      });
      await fetchList();
      setProcessedItems([]);
    } catch {
      setError("Upload failed. Please try again.");
    } finally {
      setUploading(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} className={styles.bulkUploadSpFx}>
      <Text variant="xLarge">Bulk Upload SPFx Demo</Text>
      <Text>
        Pick an Excel file to import employee records one‐by‐one or in a batch.
      </Text>

      <input
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFile}
        ref={fileInputRef}
        className={styles.fileInput}
      />

      {missingRows.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          <strong>Missing fields in rows:</strong>
          {missingRows.map((m) => (
            <div key={m.row}>
              Row {m.row}: {m.fields.join(", ")}
            </div>
          ))}
        </MessageBar>
      )}

      {processedItems.length > 0 && (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Upload Sequentially"
            onClick={() => void runUpload(false)}
          />
          <PrimaryButton
            text="Upload in Batch"
            onClick={() => void runUpload(true)}
          />
        </Stack>
      )}

      {processedItems.length > 1000 && (
        <MessageBar messageBarType={MessageBarType.info}>
          For large files (1000+ rows), batch upload is recommended.
        </MessageBar>
      )}

      {uploading && (
        <>
          <ProgressIndicator label="Uploading..." percentComplete={progress} />
          <Text>
            {sentCount} of {processedItems.length} uploaded
          </Text>
        </>
      )}

      {loading ? (
        <ProgressIndicator label="Loading existing items..." />
      ) : (
        <Text>{items.length} records in SharePoint</Text>
      )}

      {error && (
        <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
      )}
    </Stack>
  );
};

export default BulkUploadSpFx;
