# OS badges

One flat badge per OS generation, colour-graded old (dark red) -> current (green) so a diagram shows the old kit at a glance. Regenerate with `python3 tools/make_os_badges.py .`

PNG renderer detected: **rsvg-convert**.

| Badge | Means |
|---|---|
| `windows-server-2003` | WIN 2003 |
| `windows-server-2008` | WIN 2008 |
| `windows-server-2008-r2` | WIN 2008 R2 |
| `windows-server-2012` | WIN 2012 |
| `windows-server-2012-r2` | WIN 2012 R2 |
| `windows-server-2016` | WIN 2016 |
| `windows-server-2019` | WIN 2019 |
| `windows-server-2022` | WIN 2022 |
| `windows-server-2025` | WIN 2025 |
| `windows-10` | DESK 10 |
| `windows-11` | DESK 11 |
| `esxi-6.5` | ESXi 6.5 |
| `esxi-7` | ESXi 7 |
| `esxi-8` | ESXi 8 |
| `vcenter` | vCenter VC |
| `ubuntu` | Ubuntu LTS |
| `debian` | Debian 12 |
| `appliance` | appliance APP |
| `unknown` | unknown ? |

Colour scale: 2003/2008 `#5c1f1f`->`#a63d40`, 2012 R2 `#e5484d`, 2016 `#f5a524`, 2022 `#30a46c`. ESXi badges are purple, vCenter blue-purple, Linux/appliance grey or brand.
