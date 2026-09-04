class NonlinearTimerSliderCard extends HTMLElement {
  constructor() {
    super();
    this.attachShadow({ mode: "open" });

    this._config = null;
    this._hass = null;
    this._isDragging = false;
    this._values = this._buildValues();

    this._valueEl = null;
    this._sliderEl = null;
  }

  setConfig(config) {
    if (!config || !config.entity) {
      throw new Error("Entity is required");
    }

    this._config = {
      entity: config.entity,
      name: config.name || "Mute Timer (seconds)",
      icon: config.icon || "mdi:timer-sand",
    };

    this._render();
  }

  set hass(hass) {
    this._hass = hass;
    if (!this._config || !this._sliderEl || this._isDragging) {
      return;
    }

    const seconds = this._getEntitySeconds();
    const index = this._secondsToIndex(seconds);
    this._syncUi(seconds, index);
  }

  getCardSize() {
    return 2;
  }

  _buildValues() {
    const values = [];

    for (let i = 0; i <= 60; i += 1) {
      values.push(i);
    }
    for (let i = 65; i <= 120; i += 5) {
      values.push(i);
    }
    for (let i = 130; i <= 240; i += 10) {
      values.push(i);
    }

    return values;
  }

  _getEntitySeconds() {
    const stateObj =
      this._hass &&
      this._hass.states &&
      this._config &&
      this._hass.states[this._config.entity];

    const raw = stateObj ? parseInt(stateObj.state, 10) : 0;
    if (Number.isNaN(raw)) {
      return 0;
    }

    return Math.max(0, Math.min(240, raw));
  }

  _secondsToIndex(seconds) {
    let idx = 0;

    for (let i = 0; i < this._values.length; i += 1) {
      if (this._values[i] <= seconds) {
        idx = i;
      } else {
        break;
      }
    }

    return idx;
  }

  _indexToSeconds(index) {
    const safeIndex = Math.max(0, Math.min(this._values.length - 1, index));
    return this._values[safeIndex];
  }

  _setSeconds(seconds) {
    if (!this._hass || !this._config) {
      return;
    }

    this._hass.callService("input_number", "set_value", {
      entity_id: this._config.entity,
      value: seconds,
    });
  }

  _hideSliderIndicator() {
    if (!this._sliderEl) {
      return;
    }

    this._sliderEl.labeled = false;
    this._sliderEl.pin = false;
    this._sliderEl.removeAttribute("labeled");
    this._sliderEl.removeAttribute("pin");

    const sliderRoot = this._sliderEl.shadowRoot;
    if (!sliderRoot) {
      return;
    }

    if (sliderRoot.querySelector("style.nonlinear-timer-slider-hide-indicator")) {
      return;
    }

    const styleEl = document.createElement("style");
    styleEl.classList.add("nonlinear-timer-slider-hide-indicator");
    styleEl.textContent = `
      .value-indicator-container,
      .md3-slider__value-indicator-container,
      .mdc-slider__value-indicator-container,
      [part='value-indicator'] {
        display: none !important;
        opacity: 0 !important;
        visibility: hidden !important;
      }
    `;
    sliderRoot.appendChild(styleEl);
  }

  _syncUi(seconds, index) {
    if (this._valueEl) {
      this._valueEl.textContent = `${seconds}s`;
    }

    if (this._sliderEl) {
      this._sliderEl.value = index;
    }
  }

  _onInput() {
    if (!this._sliderEl) {
      return;
    }

    this._isDragging = true;

    const index = parseInt(this._sliderEl.value, 10) || 0;
    const seconds = this._indexToSeconds(index);

    this._syncUi(seconds, index);
  }

  _onChange() {
    if (!this._sliderEl) {
      return;
    }

    const index = parseInt(this._sliderEl.value, 10) || 0;
    const seconds = this._indexToSeconds(index);

    this._setSeconds(seconds);
    this._isDragging = false;
    this._syncUi(seconds, index);
  }

  _render() {
    if (!this._config) {
      return;
    }

    const liveSeconds = this._getEntitySeconds();
    const liveIndex = this._secondsToIndex(liveSeconds);

    this.shadowRoot.innerHTML = `
      <style>
        ha-card {
          padding: 10px 12px 14px;
        }

        .header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 8px;
          margin-bottom: 12px;
          color: var(--primary-text-color);
          font-size: 14px;
          line-height: 1.2;
        }

        .title {
          display: flex;
          align-items: center;
          gap: 12px;
          min-width: 0;
          font-weight: 600;
          opacity: 0.95;
        }

        .title span {
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        .value {
          flex: 0 0 auto;
          font-variant-numeric: tabular-nums;
          font-weight: 700;
          color: var(--primary-color);
        }

        .slider-wrap {
          width: 100%;
        }

        ha-slider {
          width: 100%;
          --paper-slider-secondary-color: transparent;
        }
      </style>
      <ha-card>
        <div class="header">
          <div class="title">
            <ha-icon icon="${this._config.icon}"></ha-icon>
            <span>${this._config.name}</span>
          </div>
          <span class="value">${liveSeconds}s</span>
        </div>
        <div class="slider-wrap">
          <ha-slider
            min="0"
            max="${this._values.length - 1}"
            step="1"
            value="${liveIndex}"
            ignore-bar-touch
          ></ha-slider>
        </div>
      </ha-card>
    `;

    this._valueEl = this.shadowRoot.querySelector(".value");
    this._sliderEl = this.shadowRoot.querySelector("ha-slider");

    if (!this._sliderEl) {
      return;
    }

    // Keep feedback in the top-right seconds counter only.
    this._hideSliderIndicator();
    requestAnimationFrame(() => this._hideSliderIndicator());

    const wrapper = this.shadowRoot.querySelector(".slider-wrap");
    if (wrapper) {
      wrapper.addEventListener("click", (ev) => ev.stopPropagation());
    }

    this._sliderEl.addEventListener("input", () => this._onInput());
    this._sliderEl.addEventListener("change", () => this._onChange());
  }
}

if (!customElements.get("nonlinear-timer-slider-card")) {
  customElements.define("nonlinear-timer-slider-card", NonlinearTimerSliderCard);
}

window.customCards = window.customCards || [];
window.customCards.push({
  type: "nonlinear-timer-slider-card",
  name: "Nonlinear Timer Slider Card",
  description: "Nonlinear timer slider with exact seconds display",
});
