"""
Tabs package for Heat Loss Model.
"""

from .dashboard_tab import build_dashboard_tab
from .fabric_tab import build_fabric_tab
from .dhw_pool_tab import build_dhw_pool_tab
from .systems_tab import build_systems_tab
from .room_tab import build_room_tab

__all__ = [
    "build_dashboard_tab",
    "build_fabric_tab",
    "build_dhw_pool_tab",
    "build_systems_tab",
    "build_room_tab"
]
