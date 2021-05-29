import * as dotenv from 'dotenv';
import * as path from 'path';

export default class Config {

    /**
     * Instance of Config Class
     */
    private static _instance: Config;

    /**
     * .env file data
     */
    private _env: dotenv.DotenvConfigOutput;

    /**
     * List of registered environment variable (optimization)
     */
    private _variables: Map<string, string>;

    private constructor() {
        const envFile = path.join(path.dirname(__filename), '../.env');
        this._env = dotenv.config({ path: envFile });
        this._variables = new Map;
    }

    /**
     * Get Config instance
     * @returns Config instance
     */
    public static get() {
        if (!this._instance) {
            this._instance = new Config();
        }
        return this._instance;
    }

    /**
     * Get environment varialble
     * @param s Name of the variable
     * @returns Value of the variable
     * @throws If name of variable not found
     */
    public var(s: string): string {
        if (!this._variables.get(s.toUpperCase())) {
            const v = process.env[s.toUpperCase()] || this._env.parsed[s.toUpperCase()];
            if (v) {
                this._variables.set(s.toUpperCase(), v);
            } else {
                throw new TypeError(`The environment variable '${s.toUpperCase()}' was not found.`);
            }
        }
        return this._variables.get(s.toUpperCase());
    }
}