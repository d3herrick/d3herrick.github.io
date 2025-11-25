<script type="text/javascript">
let e = document.querySelector("a[href='https://docs.google.com/forms/d/e/1FAIpQLSfY84PyDJLISzWalM7o3dFcBXjIVW-7Y3bA7hytv_SBqgZ_dA/viewform']")

if (e != null) {
  e.href += '?usp=pp_url&entry.1651182976=No&entry.1393065077=No&entry.810603448=No&entry.940799917=No';
  
  let group_name = new URLSearchParams(window.location.search).get('group_name');

  if (group_name != null) {
    e.href += '&entry.2121398355=' + group_name;
  }
}
</script>