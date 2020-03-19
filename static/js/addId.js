function addId(count) {
  document.getElementsById("updateNo").setAttribute("id", "updateNo" + count);
  document.getElementsById("masterQr").setAttribute("id", "masterQr" + count);
  document.getElementsById("readQr").setAttribute("id", "readQr" + count);
  document.getElementsById("userName").setAttribute("id", "userName" + count);
  document.getElementsById("machineName").setAttribute("id", "machineName" + count);
}
