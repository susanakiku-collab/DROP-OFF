window.DispatchExplain = {
  logs: [],

  reset() {
    this.logs = [];
  },

  log(msg) {
    this.logs.push(msg);
  },

  output() {
    return this.logs.join("\n");
  }
};