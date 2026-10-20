import { parentPort, workerData } from "node:worker_threads";

import { runLibrary, type CorpusFile, type LibraryRunInput } from "./library";

const { libraryId, input, files } = workerData as {
  libraryId: string;
  input: LibraryRunInput;
  files?: CorpusFile[];
};
const result = await runLibrary(libraryId, input, files);
parentPort?.postMessage(result);
