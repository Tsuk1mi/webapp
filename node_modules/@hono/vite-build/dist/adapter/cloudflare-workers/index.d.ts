import { Plugin } from 'vite';
import { BuildOptions } from '../../base.js';
import { GetEntryContentOptions } from '../../entry/index.js';

type CloudflareWorkersBuildOptions = BuildOptions & Pick<GetEntryContentOptions, 'entryContentAfterHooks' | 'entryContentDefaultExportHook'>;
declare const defaultOptions: CloudflareWorkersBuildOptions;
declare const cloudflareWorkersBuildPlugin: (pluginOptions?: CloudflareWorkersBuildOptions) => Plugin;

export { CloudflareWorkersBuildOptions, cloudflareWorkersBuildPlugin as default, defaultOptions };
