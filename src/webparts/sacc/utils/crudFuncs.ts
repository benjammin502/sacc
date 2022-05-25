import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";

// Get Client-List data to populate client dropdown
export const getClients = async () => {
  const response = await sp.web.lists.getByTitle('Client-List').items.select('Title').getAll();
  return response;
};

export const getItem = async (itemId) => {
  const response = await sp.web.lists.getByTitle('Client-List').items.getById(10).get();
  return response;
};