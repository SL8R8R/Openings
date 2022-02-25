function autoCaps(e) {
  if (typeof e.range.getValue() != 'object') {
  e.range.setValue(e.value.toUpperCase());
}
}
