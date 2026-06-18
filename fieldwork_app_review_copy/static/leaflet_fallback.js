(function () {
  if (window.L) return;

  function asLatLng(value) {
    if (Array.isArray(value)) return { lat: Number(value[0]) || 0, lng: Number(value[1]) || 0 };
    return { lat: Number(value.lat) || 0, lng: Number(value.lng) || 0 };
  }

  function makeBounds(points) {
    const pts = points.map(asLatLng);
    let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
    pts.forEach(p => {
      minLat = Math.min(minLat, p.lat); maxLat = Math.max(maxLat, p.lat);
      minLng = Math.min(minLng, p.lng); maxLng = Math.max(maxLng, p.lng);
    });
    if (!pts.length) minLat = maxLat = minLng = maxLng = 0;
    return {
      minLat, maxLat, minLng, maxLng,
      getCenter() { return [(minLat + maxLat) / 2, (minLng + maxLng) / 2]; },
      pad(ratio) {
        const latPad = Math.max((maxLat - minLat) * ratio, 0.01);
        const lngPad = Math.max((maxLng - minLng) * ratio, 0.01);
        return makeBounds([[minLat - latPad, minLng - lngPad], [maxLat + latPad, maxLng + lngPad]]);
      }
    };
  }

  class FallbackMap {
    constructor(id) {
      this.container = typeof id === "string" ? document.getElementById(id) : id;
      this.container.classList.add("leaflet-fallback-map");
      this.center = { lat: -29, lng: 24 };
      this.zoom = 5;
      this.layers = new Set();
      this.events = {};
      this.bounds = null;
      const badge = document.createElement("div");
      badge.className = "leaflet-fallback-badge";
      badge.textContent = "Offline map mode";
      this.container.appendChild(badge);
    }
    setView(center, zoom) { this.center = asLatLng(center); this.zoom = zoom || this.zoom; this.render(); return this; }
    getCenter() { return this.center; }
    getZoom() { return this.zoom; }
    on(name, fn) { (this.events[name] ||= []).push(fn); return this; }
    fire(name) { (this.events[name] || []).forEach(fn => fn()); }
    addLayer(layer) { this.layers.add(layer); layer._map = this; layer.render?.(); return this; }
    removeLayer(layer) { this.layers.delete(layer); layer._remove?.(); return this; }
    invalidateSize() { this.render(); return this; }
    fitBounds(bounds) { this.bounds = bounds; this.center = asLatLng(bounds.getCenter()); this.render(); return this; }
    project(latlng) {
      const p = asLatLng(latlng);
      const rect = this.container.getBoundingClientRect();
      const b = this.bounds || makeBounds([[this.center.lat - 0.08, this.center.lng - 0.08], [this.center.lat + 0.08, this.center.lng + 0.08]]);
      const lngSpan = Math.max(b.maxLng - b.minLng, 0.0001);
      const latSpan = Math.max(b.maxLat - b.minLat, 0.0001);
      return {
        x: ((p.lng - b.minLng) / lngSpan) * rect.width,
        y: ((b.maxLat - p.lat) / latSpan) * rect.height
      };
    }
    render() { this.layers.forEach(layer => layer.render?.()); }
  }

  class FeatureGroup {
    constructor() { this.layers = new Set(); this._map = null; }
    addTo(map) { this._map = map; map.addLayer(this); return this; }
    addLayer(layer) { this.layers.add(layer); if (this._map) { layer._map = this._map; layer.render?.(); } return this; }
    removeLayer(layer) { this.layers.delete(layer); layer._remove?.(); return this; }
    render() { this.layers.forEach(layer => { layer._map = this._map; layer.render?.(); }); }
    getBounds() {
      const pts = [];
      this.layers.forEach(layer => pts.push(...(layer.points?.() || [])));
      return makeBounds(pts);
    }
  }

  class Marker {
    constructor(latlng, options) { this.latlng = asLatLng(latlng); this.options = options || {}; this.events = {}; this.el = null; this.tooltip = null; }
    addTo(target) { target.addLayer ? target.addLayer(this) : target.addLayer?.(this); return this; }
    points() { return [[this.latlng.lat, this.latlng.lng]]; }
    on(name, fn) { this.events[name] = fn; return this; }
    bindTooltip(text) { this.tooltip = text; this.render(); return this; }
    getTooltip() { return this.tooltip; }
    unbindTooltip() { this.tooltip = null; this.render(); return this; }
    setIcon(icon) { this.options.icon = icon; this.render(); return this; }
    _remove() { this.el?.remove(); this.el = null; }
    render() {
      if (!this._map) return;
      if (!this.el) {
        this.el = document.createElement("div");
        this.el.className = "leaflet-fallback-marker";
        this.el.addEventListener("click", e => { e.stopPropagation(); this.events.click?.(e); });
        this._map.container.appendChild(this.el);
      }
      const icon = this.options.icon || {};
      this.el.innerHTML = icon.html || "";
      this.el.title = this.tooltip || "";
      const p = this._map.project(this.latlng);
      this.el.style.left = p.x + "px";
      this.el.style.top = p.y + "px";
    }
  }

  class Polyline {
    constructor(coords, style) { this.coords = coords.map(asLatLng); this.style = style || {}; this.events = {}; this.svg = null; this.line = null; }
    addTo(target) { target.addLayer ? target.addLayer(this) : target.addLayer?.(this); return this; }
    points() { return this.coords.map(p => [p.lat, p.lng]); }
    on(name, fn) { this.events[name] = fn; return this; }
    setStyle(style) { this.style = { ...this.style, ...style }; this.render(); return this; }
    getBounds() { return makeBounds(this.points()); }
    _remove() { this.svg?.remove(); this.svg = null; }
    render() {
      if (!this._map) return;
      if (!this.svg) {
        this.svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        this.svg.classList.add("leaflet-fallback-line");
        this.line = document.createElementNS("http://www.w3.org/2000/svg", "polyline");
        this.line.setAttribute("fill", "none");
        this.line.addEventListener("click", e => { e.stopPropagation(); this.events.click?.(e); });
        this.svg.appendChild(this.line);
        this._map.container.appendChild(this.svg);
      }
      this.line.setAttribute("points", this.coords.map(c => {
        const p = this._map.project(c);
        return `${p.x},${p.y}`;
      }).join(" "));
      this.line.setAttribute("stroke", this.style.color || "#2563eb");
      this.line.setAttribute("stroke-width", this.style.weight || 4);
      this.line.setAttribute("opacity", this.style.opacity || 1);
    }
  }

  window.L = {
    __localFallback: true,
    map: (id, options) => new FallbackMap(id, options),
    tileLayer: () => ({ addTo(map) { return map; } }),
    featureGroup: () => new FeatureGroup(),
    marker: (latlng, options) => new Marker(latlng, options),
    polyline: (coords, style) => new Polyline(coords, style),
    divIcon: options => options || {},
    DomEvent: { stopPropagation: e => e?.stopPropagation?.() }
  };
})();
