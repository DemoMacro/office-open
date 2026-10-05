import { parentPort, workerData } from "node:worker_threads";

import { runLibrary, type LibraryRunInput } from "./library";

const result = await runLibrary(workerData.libraryId, workerData.input as LibraryRunInput);
parentPort?.postMessage(result);
