import { generateUuid } from "./uuid";

export const createAppId = () => {
	return "a" + generateUuid();
};
