import { Plugin } from 'vite';
import { BuildOptions } from '../../base.js';
import { VercelBuildConfigV3, VercelServerlessFunctionConfig } from './types.js';
import '../../entry/index.js';

type VercelBuildOptions = {
    vercel?: {
        config?: VercelBuildConfigV3;
        function?: VercelServerlessFunctionConfig;
    };
} & Omit<BuildOptions, 'output' | 'outputDir'>;
declare const vercelBuildPlugin: (pluginOptions?: VercelBuildOptions) => Plugin;

export { VercelBuildOptions, vercelBuildPlugin as default };
