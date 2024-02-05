import fs from 'fs';
import chalk from 'chalk';
import ora from 'ora';

export const isDir = (filePath: string) => fs.statSync(filePath).isDirectory();

export const isFile = (filePath: string) => fs.statSync(filePath).isFile();

export const succeed = (message: string) => ora().succeed(chalk.greenBright.bold(message));

export const info = (message: string) => ora().info(chalk.blueBright.bold(message));

export const error = (message: string) => ora().fail(chalk.redBright.bold(message));
