# deployment-assets

Assets for automated deployments of software.

## icons/

Vendor logos used in customer-facing diagrams and notifications (NinjaOne, Huntress,
NetBird, Auvik, Proofpoint, Fortinet, ScalePad, Xerox).

## icons/excalidraw-libraries/

Vendored Excalidraw icon libraries (MIT) used when drawing server / network / VM
diagrams — hypervisors, hosts, blades, racks, switches, firewalls, sites, databases.
See [`icons/excalidraw-libraries/INDEX.md`](icons/excalidraw-libraries/INDEX.md) for the
full list and what lives in each one.

Refresh them with:

```bash
python3 tools/fetch_excalidraw_libraries.py .
```

### Which library for which job

| Job | Library |
|---|---|
| Hypervisors, hosts, clusters, datacenters, storage tiers, vCenter | `odraghi/vmware-architecture-design` |
| NSX-T / virtual networking | `novakkkarel/nsx-t-vmware` |
| Firewalls and Fortinet estate (FortiGate, FortiSwitch, FortiManager ...) | `fortijosh/fortinet` |
| Generic network topology (switch, router, firewall, VPN, client) | `dwelle/network-topology-icons` |
| Physical servers, racks, 1U/2U/4U blades | `jgodoy/racks-and-servers-components` |
| Sites / locations / buildings | `jgodoy/network-locations` |
| Server, database, printer, cloud, host, disk primitives | `mateuszbaransanok/it-icons` |
| Architecture components (Slack, Docker, GitHub, VPC, subnets, users) | `anna-pastushko/architecture-diagram-components` |
| Product / software logos | `maeddes/technology-logos`, `drwnio/drwnio` |
| Microsoft product icons | `zesty-lemur/microsoft-apps`, `wictorwilen/microsoft-365-icons` |

**Gap:** no library on libraries.excalidraw.com carries Windows Server edition badges
(2008 / 2012 R2 / 2016 / 2022). Those are generated in-house — see
[`icons/os-badges/`](icons/os-badges/) — so every diagram marks an OS generation the
same way.
