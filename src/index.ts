import path from 'path';
import fs from 'fs';
import ora from 'ora';
import { isDir, isFile, succeed, info, error } from './utils';
import extensions from 'text-extensions';
import ExcelJS from 'exceljs';

const desktopDir = path.resolve(__dirname, '../../../');

const dirNames = fs.readdirSync(desktopDir);

const isExcludeDir = (path: string) =>
    path.includes('node_modules') || path.includes('dist') || path.includes('inquire-command');

const commandList:string[] = [];

type Data = {
    [key: string]: {
        [key: string]: Array<string | number>;
        business: any;
    };
};

let data: Data = {};

const run = () => {
    fs.unlinkSync(path.resolve(__dirname,'../distribution.xlsx'))

    commandList.forEach((item) => (data[item] = { business: [] }));

    info(`扫描的文件有：\n${dirNames.join('\n')}`);
    const businessDirList: string[] = [];

    const businessFileList: string[] = [];

    dirNames.forEach((item) => {
        const filePath = path.join(desktopDir, item);
        isDir(filePath) ? businessDirList.push(filePath) : businessFileList.push(filePath);
    });

    const spinner: ora.Ora = ora({ text: '1.开始扫描文件...' }).start();

    businessDirList.forEach((i) => scanning(i));

    businessFileList.forEach((i) => read(i));

    spinner.stopAndPersist({ symbol: '✅', text: '1. 文件扫描完毕！' });

    output(spinner);
};

// 扫描
const scanning = (dir: string): void => {
    if (!isExcludeDir(dir)) {
        if (isFile(dir)) {
            return read(dir);
        }
        const files = fs.readdirSync(dir);
        for (let i = 0; i < files.length; i++) {
            const filePath = path.join(dir, files[i]);
            isDir(filePath) ? scanning(filePath) : read(filePath);
        }
    }
};

// 读取
const read = (filePath: string) => {
    if (!filePath.includes('.map')) {
        const list = path.extname(filePath).split('.');
        const fileExtension = extensions.includes(list.length >= 2 ? list[1] : '');
        if (fileExtension) {
            const fileContent = fs.readFileSync(filePath, 'utf-8');
            const regex = new RegExp(commandList.join('|'), 'g');
            const commandArr: string[] = fileContent.match(regex) ?? [];
            const lines = fileContent.split('\n');
            if (commandArr?.length) record(filePath, lines, commandArr);
        }
    }
};

// 记录
const record = (filePath: string, lines: string[], commandArr: string[]) => {
    for (let i = 0; i < lines.length; i++) {
        for (let j = 0; j < commandArr.length; j++) {
            const command = commandArr[j];
            if (lines[i].includes(command)) {
                const commandKeys = Object.keys(data);
                for (let l = 0; l < commandKeys.length; l++) {
                    if (commandKeys[l] === command) {
                        const pathKeys = Object.keys(data[command]);
                        if (pathKeys.includes(filePath)) {
                            data[command][filePath]?.includes(i + 1) ? '' : data[command][filePath]?.push(i + 1);
                        } else {
                            data[command] = { ...data[command], [filePath]: [i + 1] };
                        }
                    }
                    const dirs = filePath.split(path.sep);
                    const business = dirs.find(
                        (item, index) => dirs.findIndex((i) => i === 'desktop') + 2 === index,
                    );
                    if (data[command].business.length && business) {
                        if (!data[command].business.includes(business)) {
                            data[command].business.push(business);
                        }
                    } else {
                        business ? (data[command].business = [business]) : '';
                    }
                }
            }
        }
    }
};

// 输出
const output = (spinner: ora.Ora) => {
    const xlsxData = Object.keys(data).map((key: string) => {
        const locationAndRows = Object.keys(data[key])
            .splice(1)
            .map((item: string) => `${item} : ${data[key][item].join(',')}`);

        return [key, data[key].business.join('\n'), locationAndRows.join('\n')];
    });

    const workbook = new ExcelJS.Workbook();

    const worksheet: ExcelJS.Worksheet = workbook.addWorksheet('Sheet 1');

    xlsxData.unshift(['command', 'package', 'code&line']);
    xlsxData.forEach((item: string[], index) => {
        worksheet.addRow(item);
        if (index === 0) {
            worksheet.getColumn(index + 1).width = 30;
        } else if (index === 1) {
            worksheet.getColumn(index + 1).width = 50;
        } else if (index === 2) {
            worksheet.getColumn(index + 1).width = 120;
        }
    });
    for (let i = 0; i < commandList.length; i++) {
        let packages = xlsxData[i][1]?.split('\n').map((str: string) => ({ text: str })) || [
            { text: '' },
        ];

        for (let j = 1; j < packages.length; j += 2) {
            packages.splice(j, 0, { text: '\n', font: { size: 12, bold: true } });
        }
        worksheet.getCell(`B${i + 1}`).value = { richText: packages };

        let codeAndLine = xlsxData[i][2]?.split('\n').map((str: string) => ({ text: str })) || [
            { text: '' },
        ];

        for (let j = 1; j < codeAndLine.length; j += 2) {
            codeAndLine.splice(j, 0, { text: '\n', font: { size: 12, bold: true } });
        }

        worksheet.getCell(`C${i + 1}`).value = { richText: codeAndLine };

        tidy(`A${i + 1}`, worksheet);
        tidy(`B${i + 1}`, worksheet);
        tidy(`C${i + 1}`, worksheet);
    }

    workbook.xlsx
        .writeFile('distribution.xlsx')
        .then(() => {
            spinner.stopAndPersist({ symbol: '✅', text: '2. 文件已经成功生成!' });
            succeed(`文件所在位置为:${path.resolve(__dirname, '../distribution.xlsx')}`);
        })
        .catch((err) => {
            error(`导出excel失败:${err}`);
        });
};

// 整理格式
const tidy = (cell: string, worksheet: ExcelJS.Worksheet) => {
    worksheet.getCell(cell).alignment = {
        vertical: 'middle', //垂直居中

        horizontal: 'center', //水平居中

        wrapText: true, //增加自动换行属性 解决双击才会换行问题
    };
};

run();
