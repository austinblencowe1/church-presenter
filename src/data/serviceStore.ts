import { invoke, isTauri } from "@tauri-apps/api/core";
import type { Service, ServiceSummary } from "../types/presentation";

const storageKey = "church-presenter-services";

function readBrowserServices(): Service[] {
  const value = localStorage.getItem(storageKey);
  if (!value) return [];
  try {
    const parsed: unknown = JSON.parse(value);
    return Array.isArray(parsed) ? parsed as Service[] : [];
  } catch {
    return [];
  }
}

function writeBrowserServices(services: Service[]) {
  localStorage.setItem(storageKey, JSON.stringify(services));
}

export async function listServices(): Promise<ServiceSummary[]> {
  if (isTauri()) return invoke<ServiceSummary[]>("list_services");
  return readBrowserServices().map(({ id, title, date }) => ({ id, title, date }));
}

export async function loadService(serviceId: string): Promise<Service> {
  if (isTauri()) return invoke<Service>("load_service", { serviceId });
  const service = readBrowserServices().find((item) => item.id === serviceId);
  if (!service) throw new Error("Service not found");
  return service;
}

export async function saveService(service: Service): Promise<void> {
  if (isTauri()) {
    await invoke("save_service", { service });
    return;
  }
  const services = readBrowserServices();
  const index = services.findIndex((item) => item.id === service.id);
  if (index < 0) services.push(service);
  else services[index] = service;
  writeBrowserServices(services);
}

export async function deleteService(serviceId: string): Promise<void> {
  if (isTauri()) {
    await invoke("delete_service", { serviceId });
    return;
  }
  writeBrowserServices(readBrowserServices().filter((item) => item.id !== serviceId));
}