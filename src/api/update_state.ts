import { update_state_url } from "./consts"
import { get } from "./fetch"

export const update_state = async (uid: string) => {
    return await get(update_state_url + uid);
}