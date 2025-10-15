//側邊欄縮放
function toggleSidebar() {
    let sidebar = document.getElementById('sidebar');
    if(sidebar) {
       sidebar.classList.toggle('hidden');
    }
}

//側邊欄日誌選單
document.addEventListener("DOMContentLoaded", function () {
    // Use event delegation on a static parent for dynamically added elements
    document.body.addEventListener("click", function(event) {
        if (event.target.classList.contains("itemToggle")) {
            let toggle = event.target;
            event.preventDefault(); // Prevent default link behavior

            let submenu = toggle.nextElementSibling;
            let arrow = toggle.querySelector(".arrow");

            if (submenu && submenu.classList.contains("submenu")) {
                let isVisible = submenu.style.display === "block";
                
                // Close all other submenus first
                document.querySelectorAll('.submenu').forEach(function(otherSubmenu) {
                    if (otherSubmenu !== submenu) {
                       // otherSubmenu.style.display = "none";
                       // let otherArrow = otherSubmenu.previousElementSibling.querySelector(".arrow");
                       // if(otherArrow) otherArrow.innerHTML = "▼";
                    }
                });

                if (isVisible) {
                    submenu.style.display = "none";
                    if(arrow) arrow.innerHTML = "▼";
                } else {
                    submenu.style.display = "block";
                    if(arrow) arrow.innerHTML = "▲";
                }
            }
        }
    });
});