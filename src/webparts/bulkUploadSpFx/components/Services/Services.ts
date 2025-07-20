import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { ISharePointItem } from "../IBulkUploadSpFxProps";

export const getItemsFromList = async (
  context: WebPartContext,
  listName: string
): Promise<ISharePointItem[]> => {
  const sp = spfi().using(SPFx(context));
  const query = sp.web.lists
    .getByTitle(listName)
    .items.top(4999)
    .orderBy("ID", false);
  let all: ISharePointItem[] = [];
  for await (const batch of query) all = all.concat(batch);
  return all;
};

/** Sequential uploads one at a time */
export async function saveDataSequential(
  items: ISharePointItem[],
  context: WebPartContext,
  listName: string,
  onProgress?: (sent: number, total: number) => void
) {
  const sp = spfi().using(SPFx(context));
  for (let i = 0; i < items.length; i++) {
    await sp.web.lists.getByTitle(listName).items.add(items[i]);
    onProgress?.(i + 1, items.length);
  }
}

/** Batched upload */
export async function saveDataBatch(
  items: ISharePointItem[],
  context: WebPartContext,
  listName: string,
  onProgress?: (sent: number, total: number) => void
) {
  const [batchedSP, execute] = spfi().using(SPFx(context)).batched();
  let sent = 0;
  items.forEach(async (item, idx) => {
    await batchedSP.web.lists
      .getByTitle(listName)
      .items.add(item)
      .then(() => {
        sent++;
        onProgress?.(sent, items.length);
      });
  });
  await execute();
}
