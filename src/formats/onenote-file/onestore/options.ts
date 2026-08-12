/** Safety limits for malformed or hostile files. */
export interface ReaderOptions {
	maxFileNodeListFragments: number;
	maxFileNodes: number;
	maxTransactionLogFragments: number;
	maxTransactionEntries: number;
	maxObjects: number;
	maxPropertiesPerObject: number;
	maxPropertySetDepth: number;
	maxPageGraphNodes: number;
	maxInkPathValues: number;
	maxAssetBytes: number;
	maxTotalAssetBytes: number;
	strictHeaderValidation: boolean;
	validateTransactionChecksums: boolean;
}

export const DEFAULT_READER_OPTIONS: ReaderOptions = {
	maxFileNodeListFragments: 100_000,
	maxFileNodes: 2_000_000,
	maxTransactionLogFragments: 100_000,
	maxTransactionEntries: 4_000_000,
	maxObjects: 1_000_000,
	maxPropertiesPerObject: 65_536,
	maxPropertySetDepth: 128,
	maxPageGraphNodes: 100_000,
	maxInkPathValues: 1_000_000,
	// Do not discard valid user data based on an importer policy. The host's
	// actual ArrayBuffer/file limits still apply and are reported explicitly.
	maxAssetBytes: Number.MAX_SAFE_INTEGER,
	maxTotalAssetBytes: Number.MAX_SAFE_INTEGER,
	strictHeaderValidation: true,
	validateTransactionChecksums: true,
};

export interface Diagnostic {
	code: string;
	message: string;
	offset?: number;
}
