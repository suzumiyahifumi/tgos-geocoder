const { queryTgosAddressQueue } = require('./src/services/tgos');

// ===== 可改這裡 =====
const oAddresses = [
	'新竹縣尖石鄉嘉樂村２鄰麥樹仁61號',
	'新竹縣湖口鄉波羅村０１６鄰千禧路４３巷２１弄２６號二樓',
	'臺北市大安區福住里０１２鄰信義路二段１９８巷１０號',
	'新竹縣新埔鎮文山里015鄰文山路㊣頭山段２４８號',
];

queryTgosAddressQueue(oAddresses, 300)
	.then((results) => {
		for (const item of results) {
			if (item.error) {
				console.log(`${item.address} => 查詢失敗: ${item.error}`);
				continue;
			}

			const output = typeof item.result === 'string'
				? item.result
				: JSON.stringify(item.result);
			console.log(`${item.address} => ${output}`);
		}
	})
	.catch((err) => {
		console.error(err.message);
		if (err.response) {
			console.error('status:', err.response.status);
			console.error('data:', err.response.data);
		}
	});
