import excel from 'exceljs';

const PRINT_PAGE = 8;
const ROW_PER_PAGE = 41;
const COL_NAMES = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];

function generate() {
	let rand = Math.random();
	if (rand < 0.1) {
		return bracketCalc() + '＝';
	}
	if (rand < 0.2) {
		return withoutBracketCalc() + '＝';
	}
	if (rand < 0.4) {
		return floatCalc() + '＝';
	}
	return intCalc() + '＝';
}

//获取一个不为0的正整数
function getInteger(minDig: number, maxDig: number, notOne?: boolean) {
	let digit = minDig + Math.floor(Math.random() * (maxDig - minDig + 1));
	let min = Math.pow(10, digit - 1);
	if (notOne && min == 1) {
		min = 2;
	}
	let max = Math.pow(10, digit);
	return min + Math.floor(Math.random() * (max - min));
}

//2-4位正整数相加减乘除
function intCalc() {
	let rand = Math.random();
	if (rand < 0.2) {
		return intAdd();
	}
	if (rand < 0.4) {
		return intSub();
	}
	if (rand < 0.7) {
		return intMul();
	}
	return intDiv();
}

function intAdd() {
	let x = getInteger(2, 4);
	let y = getInteger(2, 4);
	return x + '＋' + y;
}

function intSub() {
	let x = getInteger(2, 4);
	let y = getInteger(2, 4);
	while (y == x) {
		y = getInteger(2, 4);
	}
	if (x > y) {
		return x + '－' + y;
	}
	return y + '－' + x;
}

function intMul() {
	let x = getInteger(2, 2);
	let y = getInteger(1, 2, true);
	return x + '×' + y;
}

function intDiv() {
	let x = getInteger(1, 3, true);
	let y = getInteger(1, 1, true);
	return (x * y) + '÷' + y;
}

//带括号的加减乘除混合运算
function bracketCalc() {
	let x = getInteger(2, 2);
	let y = getInteger(2, 2);
	let bracket = '(';
	if (Math.random() < 0.5) {	//括号内加法
		bracket += x + '＋' + y + ')';
	} else {	//括号内减法
		while (x == y) {
			y = getInteger(2, 2);
		}
		if (x < y) {
			let tmp = x;
			x = y;
			y = tmp;
		}
		bracket += x + '－' + y + ')';
	}
	if (Math.random() < 0.5) {	//括号外乘法
		let m = getInteger(1, 2, true);
		return Math.random() < 0.5 ? (m + '×' + bracket) : (bracket + '×' + m);
	}
	//括号外除法
	let m = getInteger(1, 2, true);	//得数
	let n = getInteger(1, 1, true);	//除数
	if (Math.random() < 0.5) {
		let k = m * n;	//被除数
		while (k == x) {
			x = getInteger(2, 2);
		}
		if (k < x) {
			return '(' + x + '－' + (x - k) + ')÷' + n;
		}
		return '(' + x + '＋' + (k - x) + ')÷' + n;
	}
	return (n * x) + '÷(' + (m + n) + '－' + m + ')';
}

//不带括号的加减乘除混合运算
function withoutBracketCalc() {
	if (Math.random() < 0.5) {	//搭配乘法
		let x = getInteger(2, 2);
		let y = getInteger(1, 2, true);
		let w = getInteger(2, 4);
		if (x * y > w) {
			if (Math.random() < 0.5) {
				return x + '×' + y + '＋' + w;
			}
			return x + '×' + y + '－' + w;
		}
		if (Math.random() < 0.5) {
			return w + '＋' + x + '×' + y;
		}
		return w + '－' + x + '×' + y;
	} else {	//搭配除法
		let x = getInteger(1, 3);
		let y = getInteger(1, 1, true);
		if (y == x) {
			x = getInteger(1, 3);
		}
		let w = getInteger(2, 4);
		if (x > w) {
			if (Math.random() < 0.5) {
				return (x * y) + '÷' + y + '＋' + w;
			}
			return (x * y) + '÷' + y + '－' + w;
		}
		if (Math.random() < 0.5) {
			return w + '＋' + (x * y) + '÷' + y;
		}
		return w + '－' + (x * y) + '÷' + y;
	}
}

//一位小数相加减
function floatCalc() {
	let x = getInteger(1, 3) / 10;
	let y = getInteger(1, 3) / 10;
	if (Math.random() < 0.5) {
		return x + '＋' + y;
	}
	while (x == y) {
		y = getInteger(1, 3) / 10;
	}
	if (x > y) {
		return x + '－' + y;
	}
	return y + '－' + x;
}

function main() {
	const workbook = new excel.Workbook();
	let sheet = workbook.addWorksheet('Sheet1', {
		pageSetup: {
			margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
		}
	});

	for (let page = 0; page < PRINT_PAGE; ++page) {
		let start = ROW_PER_PAGE * page;
		for (let i = 1; i <= ROW_PER_PAGE; ++i) {
			for (let j = 0; j < COL_NAMES.length; j++) {
				let key = COL_NAMES[j] + (start + i);
				let cell = sheet.getCell(key);
				cell.font = { size: 12, name: '微软雅黑' };
				if (j % 4 == 0) {
					cell.value = generate();
					console.log(key, cell.value);
				}
			}
		}
	}
	workbook.xlsx.writeFile('小学三年级计算.xlsx');
}
main();
