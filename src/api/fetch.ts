import axios from 'axios'
import { base_url } from './consts';

export const get = async (url: string) => {
    return (await axios.get(base_url + url)).data;
}

export const post = async (url: string, data: any) => {
    return (await axios.post(base_url+url, data=data)).data;
}
