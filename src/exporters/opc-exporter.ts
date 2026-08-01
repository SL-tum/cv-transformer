import type { NativeStore } from "../model/core/native.js";
import { loadOpcPackage, nativeStoreToPackage, packageToNativeStore, writeOpcPackage } from "../ooxml/opc/package.js";

export const unpackOfficePackage = (input: Uint8Array): NativeStore => packageToNativeStore(loadOpcPackage(input));
export const repackOfficePackage = (store: NativeStore): Uint8Array => writeOpcPackage(nativeStoreToPackage(store));
