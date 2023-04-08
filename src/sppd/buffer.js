import axios from 'axios'
import { Buffer } from 'buffer'

export async function getResponseAsBuffer(url) {
  try {
    const response = await axios.get(url, { responseType: 'arraybuffer' });
    const buffer = Buffer.from(response.data, 'binary');
    return buffer;
  } catch (error) {
    console.error(error);
    return null;
  }
}

