/**
 * Read-only Compound File Binary (CFB v3) access.
 *
 * @module
 */

const HEADER_SIZE = 512;
const SECTOR_SHIFT = 9;
const SECTOR_SIZE = 1 << SECTOR_SHIFT;
const MINI_SECTOR_SHIFT = 6;
const MINI_SECTOR_SIZE = 1 << MINI_SECTOR_SHIFT;
const DIRECTORY_ENTRY_SIZE = 128;
const END_OF_CHAIN = 0xfffffffe;
const FREE_SECT = 0xffffffff;
const FAT_SECT = 0xfffffffd;
const DIF_SECT = 0xfffffffc;
const NO_STREAM = 0xffffffff;
const STREAM_CUTOFF = 4096;
const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

export type CompoundFileEntryType = "storage" | "stream";

/** One non-root CFB directory entry. Stream size is in bytes (0 = empty). */
export interface CompoundFileEntry {
  /** Entry name without parent directories, preserving source casing. */
  readonly name: string;
  /** Slash-separated path from the root storage, excluding the root name. */
  readonly path: string;
  readonly type: CompoundFileEntryType;
  /** Stream payload size in bytes; storages are always 0. */
  readonly size: number;
}

interface Header {
  readonly fatSectorCount: number;
  readonly directorySector: number;
  readonly miniFatSector: number;
  readonly miniFatSectorCount: number;
  readonly difatSector: number;
  readonly difatSectorCount: number;
}

interface DirectoryEntry {
  readonly id: number;
  readonly name: string;
  readonly type: CompoundFileEntryType;
  readonly size: number;
  readonly leftSiblingId: number;
  readonly childId: number;
  readonly rightSiblingId: number;
  readonly startSector: number;
  path?: string;
}

interface RootEntry {
  readonly childId: number;
  readonly startSector: number;
  readonly size: number;
}

interface Contents {
  readonly fat: Uint32Array;
  readonly root: RootEntry;
  entries: readonly CompoundFileEntry[];
  readonly paths: Map<string, DirectoryEntry>;
}

/** Random-access reader for an in-memory CFB v3 container. */
export class CompoundFileReader {
  private readonly bytes: Uint8Array;
  private readonly view: DataView;
  private readonly header: Header;
  private readonly contents: Contents;
  private readonly directory: Map<number, DirectoryEntry>;
  private miniFat?: Uint32Array;
  private miniStream?: Uint8Array;

  public constructor(data: Uint8Array) {
    this.bytes = data;
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const header = this.readHeader();
    this.header = header;

    const fat = this.readFat();
    const directorySectors = this.followChain(header.directorySector, "directory", fat, fat.length);
    this.directory = this.parseDirectory(this.concatSectors(directorySectors));
    this.contents = {
      fat,
      root: this.readRoot(),
      entries: [],
      paths: new Map(),
    };
    this.buildDirectoryTree();
    this.contents.entries = Object.freeze(this.contents.entries);
  }

  /** Immutable directory entries in directory-tree order, excluding root. */
  public get entries(): readonly CompoundFileEntry[] {
    return this.contents.entries;
  }

  /** Return an entry, or undefined when no CFB entry has that exact path. */
  public entry(path: string): CompoundFileEntry | undefined {
    const entry = this.contents.paths.get(path.toUpperCase());
    return entry?.path
      ? { name: entry.name, path: entry.path, type: entry.type, size: entry.size }
      : undefined;
  }

  /** Copy stream bytes; each call returns a new buffer owned by the caller. */
  public read(path: string): Uint8Array {
    const entry = this.contents.paths.get(path.toUpperCase());
    if (!entry?.path) throw new RangeError(`CFB stream not found: ${path}`);
    if (entry.type !== "stream") throw new TypeError(`CFB entry is not a stream: ${path}`);
    if (entry.size === 0) return new Uint8Array();
    return entry.size < STREAM_CUTOFF ? this.readMiniStream(entry) : this.readRegularStream(entry);
  }

  private readHeader(): Header {
    if (this.bytes.byteLength < HEADER_SIZE)
      throw new Error("Invalid CFB file: header is truncated");
    for (const [index, byte] of SIGNATURE.entries()) {
      if (this.bytes[index] !== byte) throw new Error("Invalid CFB file: bad signature");
    }

    const majorVersion = this.view.getUint16(26, true);
    const byteOrder = this.view.getUint16(28, true);
    const sectorShift = this.view.getUint16(30, true);
    const miniSectorShift = this.view.getUint16(32, true);
    const directorySectorCount = this.view.getUint32(40, true);
    const fatSectorCount = this.view.getUint32(44, true);
    const directorySector = this.view.getUint32(48, true);
    const miniStreamCutoff = this.view.getUint32(56, true);
    const miniFatSector = this.view.getUint32(60, true);
    const miniFatSectorCount = this.view.getUint32(64, true);
    const difatSector = this.view.getUint32(68, true);
    const difatSectorCount = this.view.getUint32(72, true);

    if (majorVersion !== 3) throw new Error(`Unsupported CFB version: ${majorVersion}`);
    if (byteOrder !== 0xfffe) throw new Error("Invalid CFB file: byte order is not little-endian");
    if (sectorShift !== SECTOR_SHIFT)
      throw new Error(`Invalid CFB v3 sector shift: ${sectorShift}`);
    if (miniSectorShift !== MINI_SECTOR_SHIFT)
      throw new Error(`Invalid CFB mini-sector shift: ${miniSectorShift}`);
    if (directorySectorCount !== 0) throw new Error("Invalid CFB v3 directory-sector count");
    if (miniStreamCutoff !== STREAM_CUTOFF)
      throw new Error(`Invalid CFB mini-stream cutoff: ${miniStreamCutoff}`);
    if (fatSectorCount === 0 || fatSectorCount > this.sectorCount) {
      throw new Error("Invalid CFB file: FAT-sector count is out of bounds");
    }
    if (difatSectorCount > this.sectorCount) {
      throw new Error("Invalid CFB file: DIFAT-sector count is out of bounds");
    }

    return {
      fatSectorCount,
      directorySector,
      miniFatSector,
      miniFatSectorCount,
      difatSector,
      difatSectorCount,
    };
  }

  private get sectorCount(): number {
    return Math.floor((this.bytes.byteLength - HEADER_SIZE) / SECTOR_SIZE);
  }

  private readFat(): Uint32Array {
    const sectorCount = this.sectorCount;
    const expectedDifatSectors = Math.max(0, Math.ceil((this.header.fatSectorCount - 109) / 127));
    if (this.header.difatSectorCount !== expectedDifatSectors) {
      throw new Error("Invalid CFB file: DIFAT-sector count does not match FAT size");
    }
    if (expectedDifatSectors === 0 && this.header.difatSector !== END_OF_CHAIN) {
      throw new Error("Invalid CFB file: DIFAT chain is present unexpectedly");
    }

    const fatSectors: number[] = [];
    for (let index = 0; index < 109; index++) {
      const sector = this.view.getUint32(76 + index * 4, true);
      if (sector !== FREE_SECT) fatSectors.push(sector);
    }

    if (expectedDifatSectors > 0) {
      let sector = this.header.difatSector;
      const seen = new Set<number>();
      for (let index = 0; index < expectedDifatSectors; index++) {
        const bytes = this.readSector(sector, "DIFAT");
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        if (seen.has(sector)) throw new Error("Invalid CFB file: DIFAT chain contains a cycle");
        seen.add(sector);
        for (let offset = 0; offset < SECTOR_SIZE - 4; offset += 4) {
          const fatSector = view.getUint32(offset, true);
          if (fatSector !== FREE_SECT) fatSectors.push(fatSector);
        }

        sector = view.getUint32(SECTOR_SIZE - 4, true);
        const isLast = index + 1 === expectedDifatSectors;
        if (isLast ? sector !== END_OF_CHAIN : sector === END_OF_CHAIN) {
          throw new Error("Invalid CFB file: DIFAT chain length is invalid");
        }
        if (sector !== END_OF_CHAIN && sector >= sectorCount) {
          throw new RangeError(`Invalid CFB DIFAT sector: ${sector}`);
        }
      }
    }

    if (fatSectors.length !== this.header.fatSectorCount) {
      throw new Error(
        `Invalid CFB file: DIFAT has ${fatSectors.length}/${this.header.fatSectorCount} FAT sectors`,
      );
    }

    const fat = new Uint32Array(sectorCount).fill(FREE_SECT);
    const seen = new Set<number>();
    for (const [fatSectorIndex, sector] of fatSectors.entries()) {
      if (sector >= sectorCount) throw new RangeError(`Invalid CFB FAT sector: ${sector}`);
      if (seen.has(sector)) throw new Error("Invalid CFB file: FAT sector is referenced twice");
      seen.add(sector);
      const bytes = this.readSector(sector, "FAT");
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      const firstEntry = fatSectorIndex * 128;
      for (let index = 0; index < 128; index++) {
        fat[firstEntry + index] = view.getUint32(index * 4, true);
      }
    }
    return fat;
  }

  private readSector(sector: number, purpose: string): Uint8Array {
    if (!Number.isInteger(sector) || sector < 0 || sector >= this.sectorCount) {
      throw new RangeError(`Invalid CFB ${purpose} sector: ${sector}`);
    }
    const offset = HEADER_SIZE + sector * SECTOR_SIZE;
    return this.bytes.slice(offset, offset + SECTOR_SIZE);
  }

  private followChain(
    start: number,
    purpose: string,
    fat: Uint32Array,
    maximumLength: number,
  ): number[] {
    if (start === END_OF_CHAIN) return [];
    if (start >= fat.length) throw new RangeError(`Invalid CFB ${purpose} start sector: ${start}`);

    const sectors: number[] = [];
    const seen = new Set<number>();
    let sector = start;
    while (true) {
      if (sector >= fat.length) throw new RangeError(`Invalid CFB ${purpose} sector: ${sector}`);
      if (seen.has(sector)) throw new Error(`Invalid CFB file: ${purpose} chain contains a cycle`);
      this.readSector(sector, purpose);
      seen.add(sector);
      sectors.push(sector);
      if (sectors.length > maximumLength)
        throw new Error(`Invalid CFB file: ${purpose} chain is too long`);

      const next = fat[sector]!;
      if (next === END_OF_CHAIN) return sectors;
      if (next === FREE_SECT || next === FAT_SECT || next === DIF_SECT) {
        throw new Error(`Invalid CFB file: ${purpose} chain has an invalid link`);
      }
      if (next >= fat.length) throw new RangeError(`Invalid CFB ${purpose} sector: ${next}`);
      sector = next;
    }
  }

  private concatSectors(sectors: readonly number[]): Uint8Array {
    const bytes = new Uint8Array(sectors.length * SECTOR_SIZE);
    sectors.forEach((sector, index) => {
      const offset = HEADER_SIZE + sector * SECTOR_SIZE;
      bytes.set(this.bytes.subarray(offset, offset + SECTOR_SIZE), index * SECTOR_SIZE);
    });
    return bytes;
  }

  private parseDirectory(bytes: Uint8Array): Map<number, DirectoryEntry> {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const directory = new Map<number, DirectoryEntry>();
    for (let id = 0; id < bytes.byteLength / DIRECTORY_ENTRY_SIZE; id++) {
      const offset = id * DIRECTORY_ENTRY_SIZE;
      const objectType = bytes[offset + 66];
      if (objectType === 0) continue;
      if (objectType !== 1 && objectType !== 2 && objectType !== 5) {
        throw new Error(`Invalid CFB directory object type: ${objectType}`);
      }

      const name = id === 0 ? "Root Entry" : this.readName(bytes, offset);
      const leftSiblingId = view.getUint32(offset + 68, true);
      const rightSiblingId = view.getUint32(offset + 72, true);
      const childId = view.getUint32(offset + 76, true);
      const startSector = view.getUint32(offset + 116, true);
      const size = view.getUint32(offset + 120, true);
      if (objectType === 5 && id !== 0) {
        throw new Error("Invalid CFB file: root entry is not directory entry zero");
      }

      directory.set(id, {
        id,
        name,
        type: objectType === 2 ? "stream" : "storage",
        size,
        leftSiblingId,
        childId,
        rightSiblingId,
        startSector,
      });
    }

    const root = directory.get(0);
    if (!root || root.name.length === 0)
      throw new Error(`Invalid CFB file: root entry is missing (${directory.size})`);
    return directory;
  }

  private readName(bytes: Uint8Array, offset: number): string {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const byteLength = view.getUint16(offset + 64, true);
    if (byteLength < 2 || byteLength > 64 || byteLength % 2 !== 0) {
      throw new Error(`Invalid CFB directory name length: ${byteLength}`);
    }

    let name: string;
    try {
      name = new TextDecoder("utf-16le", { fatal: true }).decode(
        bytes.subarray(offset, offset + byteLength - 2),
      );
    } catch {
      throw new Error("Invalid CFB directory name encoding");
    }
    if (view.getUint16(offset + byteLength - 2, true) !== 0 || name.includes("\0")) {
      throw new Error("Invalid CFB directory name terminator");
    }
    if (name.length * 2 + 2 !== byteLength || name === "." || name === "..") {
      throw new Error("Invalid CFB directory name");
    }
    if (name.includes("/") || name.includes("\\") || name.includes(":") || name.includes("!")) {
      throw new Error(`Invalid CFB directory name: ${name}`);
    }
    return name;
  }

  private readRoot(): RootEntry {
    const root = this.directory.get(0);
    if (!root || root.type !== "storage")
      throw new Error("Invalid CFB file: root entry is missing");
    return { childId: root.childId, startSector: root.startSector, size: root.size };
  }

  private buildDirectoryTree(): void {
    const { root, paths } = this.contents;
    const entries: CompoundFileEntry[] = [];
    const visited = new Set<number>([0]);

    const pending = [{ id: root.childId, parentPath: "" }];
    while (pending.length > 0) {
      const { id, parentPath } = pending.pop()!;
      if (id === NO_STREAM) continue;
      const entry = this.directory.get(id);
      if (!entry) throw new RangeError(`Invalid CFB directory entry reference: ${id}`);
      if (visited.has(id)) throw new Error("Invalid CFB file: directory tree contains a cycle");
      visited.add(id);

      const path = parentPath ? `${parentPath}/${entry.name}` : entry.name;
      entry.path = path;
      entries.push(Object.freeze({ name: entry.name, path, type: entry.type, size: entry.size }));
      if (paths.has(path.toUpperCase())) throw new Error(`Duplicate CFB path: ${path}`);
      paths.set(path.toUpperCase(), entry);
      if (entry.type === "stream" && entry.childId !== NO_STREAM) {
        throw new Error(`Invalid CFB file: stream has children: ${path}`);
      }

      pending.push(
        { id: entry.rightSiblingId, parentPath },
        { id: entry.leftSiblingId, parentPath },
        { id: entry.childId, parentPath: path },
      );
    }
    for (const [id] of this.directory) {
      if (id !== 0 && !visited.has(id)) {
        throw new Error("Invalid CFB file: directory contains unreachable entries");
      }
    }
    this.contents.entries = Object.freeze(entries);
  }

  private readRegularStream(entry: DirectoryEntry): Uint8Array {
    const expectedSectors = Math.ceil(entry.size / SECTOR_SIZE);
    const sectors = this.followChain(
      entry.startSector,
      "stream",
      this.contents.fat,
      expectedSectors,
    );
    if (sectors.length !== expectedSectors) {
      throw new Error("Invalid CFB file: stream-sector chain is incomplete");
    }

    const bytes = new Uint8Array(entry.size);
    let offset = 0;
    for (const sector of sectors) {
      const length = Math.min(SECTOR_SIZE, entry.size - offset);
      const sourceOffset = HEADER_SIZE + sector * SECTOR_SIZE;
      bytes.set(this.bytes.subarray(sourceOffset, sourceOffset + length), offset);
      offset += length;
    }
    return bytes;
  }

  private readMiniStream(entry: DirectoryEntry): Uint8Array {
    const miniFat = this.readMiniFat();
    const root = this.readRootMiniStream();
    const expectedSectors = Math.ceil(entry.size / MINI_SECTOR_SIZE);
    if (entry.startSector === END_OF_CHAIN) {
      throw new Error("Invalid CFB file: mini-stream chain is missing");
    }

    let sector = entry.startSector;
    const seen = new Set<number>();
    const bytes = new Uint8Array(entry.size);
    let offset = 0;
    while (true) {
      if (sector >= miniFat.length)
        throw new RangeError(`Invalid CFB mini-stream sector: ${sector}`);
      if (seen.has(sector)) throw new Error("Invalid CFB file: mini-stream chain contains a cycle");
      if (sector * MINI_SECTOR_SIZE + MINI_SECTOR_SIZE > root.byteLength) {
        throw new RangeError(`Invalid CFB mini-stream sector: ${sector}`);
      }
      seen.add(sector);
      if (seen.size > expectedSectors)
        throw new Error("Invalid CFB file: mini-stream chain is too long");

      const length = Math.min(MINI_SECTOR_SIZE, entry.size - offset);
      const sourceOffset = sector * MINI_SECTOR_SIZE;
      bytes.set(root.subarray(sourceOffset, sourceOffset + length), offset);
      offset += length;
      if (offset === entry.size) return bytes;

      const next = miniFat[sector]!;
      if (next === END_OF_CHAIN)
        throw new Error("Invalid CFB file: mini-stream chain is incomplete");
      if (next === FREE_SECT || next === FAT_SECT || next === DIF_SECT) {
        throw new Error("Invalid CFB file: mini-stream chain has an invalid link");
      }
      sector = next;
    }
  }

  private readMiniFat(): Uint32Array {
    if (this.miniFat) return this.miniFat;
    if (this.header.miniFatSectorCount === 0)
      throw new Error("Invalid CFB file: Mini FAT is missing");
    const sectors = this.followChain(
      this.header.miniFatSector,
      "Mini FAT",
      this.contents.fat,
      this.header.miniFatSectorCount,
    );
    if (sectors.length !== this.header.miniFatSectorCount) {
      throw new Error("Invalid CFB file: Mini FAT-sector chain is incomplete");
    }

    const miniFat = new Uint32Array(sectors.length * 128).fill(FREE_SECT);
    sectors.forEach((sector, sectorIndex) => {
      const sourceOffset = HEADER_SIZE + sector * SECTOR_SIZE;
      for (let index = 0; index < 128; index++) {
        miniFat[sectorIndex * 128 + index] = this.view.getUint32(sourceOffset + index * 4, true);
      }
    });
    this.miniFat = miniFat;
    return miniFat;
  }

  private readRootMiniStream(): Uint8Array {
    if (this.miniStream) return this.miniStream;
    const root = this.contents.root;
    const expectedSectors = Math.ceil(root.size / SECTOR_SIZE);
    const sectors = this.followChain(
      root.startSector,
      "root mini stream",
      this.contents.fat,
      expectedSectors,
    );
    if (sectors.length !== expectedSectors) {
      throw new Error("Invalid CFB file: root mini-stream chain is incomplete");
    }

    const bytes = new Uint8Array(root.size);
    let offset = 0;
    for (const sector of sectors) {
      const length = Math.min(SECTOR_SIZE, root.size - offset);
      const sourceOffset = HEADER_SIZE + sector * SECTOR_SIZE;
      bytes.set(this.bytes.subarray(sourceOffset, sourceOffset + length), offset);
      offset += length;
    }
    this.miniStream = bytes;
    return bytes;
  }
}
