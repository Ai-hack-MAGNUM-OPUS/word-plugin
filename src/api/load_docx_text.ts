import { load_docx_text_url } from "./consts"
import { post } from "./fetch"

export const load_docx = async (text: string) => {
    return await post(load_docx_text_url, {text: JSON.stringify(text)})
}
