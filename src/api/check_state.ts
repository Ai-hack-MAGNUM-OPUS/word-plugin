import { check_state_url } from "./consts"
import { get } from "./fetch"

export const check_state = async (uuid: string) => {
    return await get(check_state_url + uuid);
}