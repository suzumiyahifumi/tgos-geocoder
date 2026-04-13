const datasetLabelElement = document.getElementById('mapDatasetLabel');
const markerCountElement = document.getElementById('mapMarkerCount');
const emptyStateElement = document.getElementById('mapEmptyState');
const popupElement = document.getElementById('popup');

const { Map, View, Feature, Overlay } = ol;
const { Point } = ol.geom;
const { Tile: TileLayer, Vector: VectorLayer } = ol.layer;
const { OSM, Vector: VectorSource } = ol.source;
const { boundingExtent } = ol.extent;
const { register } = ol.proj.proj4;
const { fromLonLat, get: getProjection, transform } = ol.proj;
const { Circle: CircleStyle, Fill, Stroke, Style } = ol.style;

proj4.defs('EPSG:3826', '+proj=tmerc +lat_0=0 +lon_0=121 +k=0.9999 +x_0=250000 +y_0=0 +ellps=GRS80 +units=m +no_defs +type=crs');
proj4.defs('EPSG:3825', '+proj=tmerc +lat_0=0 +lon_0=119 +k=0.9999 +x_0=250000 +y_0=0 +ellps=GRS80 +units=m +no_defs +type=crs');
register(proj4);

let selectedPopupRowId = '';
let hoveredMarkerRowId = '';

const markerStyles = {
	default: new Style({
		image: new CircleStyle({
			radius: 8,
			fill: new Fill({ color: '#b6552f' }),
			stroke: new Stroke({ color: '#fff8ef', width: 3 }),
		}),
	}),
	hover: new Style({
		image: new CircleStyle({
			radius: 10,
			fill: new Fill({ color: '#dc7b46' }),
			stroke: new Stroke({ color: '#fff8ef', width: 4 }),
		}),
	}),
	selected: new Style({
		image: new CircleStyle({
			radius: 11,
			fill: new Fill({ color: '#2f7d6a' }),
			stroke: new Stroke({ color: '#f4fffb', width: 4 }),
		}),
	}),
};

const vectorSource = new VectorSource();
const vectorLayer = new VectorLayer({
	source: vectorSource,
	style: (feature) => {
		const rowId = feature.getId();
		if (rowId && rowId === selectedPopupRowId) {
			return markerStyles.selected;
		}

		if (rowId && rowId === hoveredMarkerRowId) {
			return markerStyles.hover;
		}

		return markerStyles.default;
	},
});

const popupOverlay = new Overlay({
	element: popupElement,
	offset: [0, -18],
	positioning: 'bottom-center',
	stopEvent: true,
});

const map = new Map({
	target: 'map',
	layers: [
		new TileLayer({
			source: new OSM(),
		}),
		vectorLayer,
	],
	overlays: [popupOverlay],
	view: new View({
		center: fromLonLat([121.5654, 25.033]),
		zoom: 12,
	}),
});

let latestPayload = {
	datasetId: '',
	datasetLabel: '',
	markers: [],
};
let lastFittedSignature = '';

function escapeHtml(value) {
	return String(value)
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

function parseNumber(value) {
	const normalized = String(value ?? '').replaceAll(',', '').trim();
	if (!normalized) {
		return null;
	}

	const parsed = Number(normalized);
	return Number.isFinite(parsed) ? parsed : null;
}

function inferSourceProjection(coordSystem, x, y) {
	const normalized = String(coordSystem || '').toLowerCase();

	if (normalized.includes('4326') || normalized.includes('wgs84') || normalized.includes('lng') || normalized.includes('lat')) {
		return 'EPSG:4326';
	}

	if (normalized.includes('3826') || normalized.includes('twd97') || normalized.includes('tm2') || normalized.includes('121')) {
		return 'EPSG:3826';
	}

	if (normalized.includes('3825') || normalized.includes('119')) {
		return 'EPSG:3825';
	}

	if (Math.abs(x) <= 180 && Math.abs(y) <= 90) {
		return 'EPSG:4326';
	}

	if (x >= 80000 && x <= 450000 && y >= 2300000 && y <= 2900000) {
		return 'EPSG:3826';
	}

	return null;
}

function toMapCoordinate(marker) {
	const x = parseNumber(marker.coordX);
	const y = parseNumber(marker.coordY);
	if (x === null || y === null) {
		return null;
	}

	const sourceProjection = inferSourceProjection(marker.coordSystem, x, y);
	if (!sourceProjection) {
		return null;
	}

	try {
		return transform([x, y], sourceProjection, 'EPSG:3857');
	} catch (_error) {
		return null;
	}
}

function buildPopupRows(marker) {
	const preferredKeys = [
		['列號', String(marker.rowNumber || '')],
		['候選地址', marker.candidateAddress || ''],
		['匹配地址', marker.matchedAddress || ''],
		['座標系統', marker.coordSystem || ''],
		['X', marker.coordX || ''],
		['Y', marker.coordY || ''],
	];

	const rowData = marker.rowData && typeof marker.rowData === 'object' ? marker.rowData : {};
	const extraRows = Object.entries(rowData)
		.filter(([key, value]) => key !== '__rowId' && !['candidate_address', 'matched_address', 'coord_x', 'coord_y', 'coord_system'].includes(key) && String(value ?? '').trim() !== '')
		.slice(0, 6)
		.map(([key, value]) => [key, String(value)]);

	return [...preferredKeys, ...extraRows]
		.filter(([, value]) => String(value ?? '').trim() !== '')
		.map(([label, value]) => `
			<div class="popup-row">
				<strong>${escapeHtml(label)}</strong>
				<span>${escapeHtml(value)}</span>
			</div>
		`)
		.join('');
}

function renderPopup(marker, coordinate) {
	popupElement.innerHTML = `
		<div class="popup-header">
			<h2 class="popup-title">第 ${escapeHtml(String(marker.rowNumber || ''))} 列</h2>
			<button class="popup-close-button" type="button" aria-label="關閉資訊視窗">×</button>
		</div>
		<div class="popup-list">${buildPopupRows(marker)}</div>
	`;
	popupElement.hidden = false;
	popupOverlay.setPosition(coordinate);
	vectorLayer.changed();
	popupElement.querySelector('.popup-close-button')?.addEventListener('click', (event) => {
		event.preventDefault();
		event.stopPropagation();
		selectedPopupRowId = '';
		clearPopup();
	});
}

function clearPopup() {
	popupElement.hidden = true;
	popupOverlay.setPosition(undefined);
	vectorLayer.changed();
}

function fitMarkers(coordinates) {
	if (coordinates.length === 0) {
		map.getView().setCenter(fromLonLat([121.5654, 25.033]));
		map.getView().setZoom(12);
		lastFittedSignature = '';
		return;
	}

	if (coordinates.length === 1) {
		map.getView().setCenter(coordinates[0]);
		map.getView().setZoom(16);
		return;
	}

	map.getView().fit(boundingExtent(coordinates), {
		padding: [60, 60, 60, 60],
		maxZoom: 17,
		duration: 250,
	});
}

function renderPayload(payload) {
	latestPayload = payload && typeof payload === 'object'
		? payload
		: {
			datasetId: '',
			datasetLabel: '',
			markers: [],
		};

	const markers = Array.isArray(latestPayload.markers) ? latestPayload.markers : [];
	const features = [];
	const coordinates = [];
	const nextSignature = JSON.stringify(markers.map((marker) => ({
		rowId: marker.rowId,
		coordX: marker.coordX,
		coordY: marker.coordY,
		coordSystem: marker.coordSystem,
	})));

	for (const marker of markers) {
		const coordinate = toMapCoordinate(marker);
		if (!coordinate) {
			continue;
		}

		const feature = new Feature({
			geometry: new Point(coordinate),
			marker,
		});
		feature.setId(marker.rowId);
		features.push(feature);
		coordinates.push(coordinate);
	}

	vectorSource.clear(true);
	if (features.length > 0) {
		vectorSource.addFeatures(features);
	}

	datasetLabelElement.textContent = latestPayload.datasetLabel
		? `${latestPayload.datasetLabel}，目前顯示勾選且已有座標的列。`
		: '尚未指定資料分頁，地圖已就緒。';
	markerCountElement.textContent = `${features.length} 個點位`;
	emptyStateElement.classList.toggle('hidden', features.length > 0);
	if (nextSignature !== lastFittedSignature) {
		fitMarkers(coordinates);
		lastFittedSignature = nextSignature;
	}

	if (selectedPopupRowId) {
		const matchedFeature = features.find((feature) => feature.getId() === selectedPopupRowId);
		if (matchedFeature) {
			const marker = matchedFeature.get('marker');
			const geometry = matchedFeature.getGeometry();
			const coordinate = geometry && typeof geometry.getCoordinates === 'function'
				? geometry.getCoordinates()
				: null;
			if (marker && coordinate) {
				renderPopup(marker, coordinate);
			}
		} else {
			selectedPopupRowId = '';
			clearPopup();
		}
	} else {
		clearPopup();
	}

	window.setTimeout(() => {
		map.updateSize();
	}, 0);
}

map.on('singleclick', async (event) => {
	const feature = map.forEachFeatureAtPixel(event.pixel, (item) => item);
	if (!feature) {
		selectedPopupRowId = '';
		clearPopup();
		return;
	}

	const marker = feature.get('marker');
	const geometry = feature.getGeometry();
	const coordinate = geometry && typeof geometry.getCoordinates === 'function'
		? geometry.getCoordinates()
		: null;
	if (!marker || !coordinate) {
		selectedPopupRowId = '';
		clearPopup();
		return;
	}

	selectedPopupRowId = marker.rowId || '';
	renderPopup(marker, coordinate);
	await window.desktopApi.focusRowFromMapEditor({
		datasetId: latestPayload.datasetId,
		rowId: marker.rowId,
	});
});

map.on('pointermove', (event) => {
	if (event.dragging) {
		return;
	}

	const feature = map.forEachFeatureAtPixel(event.pixel, (item) => item);
	const nextHoveredRowId = feature?.getId?.() || '';
	if (nextHoveredRowId === hoveredMarkerRowId) {
		return;
	}

	hoveredMarkerRowId = nextHoveredRowId;
	map.getTargetElement().style.cursor = hoveredMarkerRowId ? 'pointer' : '';
	vectorLayer.changed();
});

window.addEventListener('resize', () => {
	map.updateSize();
});

popupElement.addEventListener('click', (event) => {
	event.stopPropagation();
});

window.desktopApi.onMapEditorData((payload) => {
	renderPayload(payload);
});

if (!getProjection('EPSG:3826') || !getProjection('EPSG:3825')) {
	console.warn('TWD97 投影註冊失敗，部分座標可能無法正確轉換。');
}

map.updateSize();
renderPayload(latestPayload);
