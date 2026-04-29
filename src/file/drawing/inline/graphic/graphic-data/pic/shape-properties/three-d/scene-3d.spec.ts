import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createScene3D } from "./scene-3d";

describe("createScene3D", () => {
    it("should create a scene3d with camera and lightRig", () => {
        const tree = new Formatter().format(
            createScene3D({
                camera: { preset: "perspectiveFront" },
                lightRig: { rig: "threePt", direction: "t" },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scene3d": [
                {
                    "a:camera": { _attr: { prst: "perspectiveFront" } },
                },
                {
                    "a:lightRig": { _attr: { rig: "threePt", dir: "t" } },
                },
            ],
        });
    });

    it("should create a scene3d with camera rotation", () => {
        const tree = new Formatter().format(
            createScene3D({
                camera: {
                    preset: "isometricTopUp",
                    rotation: { lat: 0, lon: 0, rev: 5400000 },
                },
                lightRig: { rig: "balanced", direction: "tl" },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scene3d": [
                {
                    "a:camera": [
                        { _attr: { prst: "isometricTopUp" } },
                        { "a:rot": { _attr: { lat: 0, lon: 0, rev: 5400000 } } },
                    ],
                },
                {
                    "a:lightRig": { _attr: { rig: "balanced", dir: "tl" } },
                },
            ],
        });
    });

    it("should create a scene3d with camera fov and zoom", () => {
        const tree = new Formatter().format(
            createScene3D({
                camera: {
                    preset: "perspectiveFront",
                    fov: 5400000,
                    zoom: "150%",
                },
                lightRig: { rig: "threePt", direction: "t" },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scene3d": [
                {
                    "a:camera": { _attr: { prst: "perspectiveFront", fov: 5400000, zoom: "150%" } },
                },
                {
                    "a:lightRig": { _attr: { rig: "threePt", dir: "t" } },
                },
            ],
        });
    });

    it("should create a scene3d with lightRig rotation", () => {
        const tree = new Formatter().format(
            createScene3D({
                camera: { preset: "perspectiveFront" },
                lightRig: {
                    rig: "soft",
                    direction: "r",
                    rotation: { lat: 0, lon: 0, rev: 2700000 },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scene3d": [
                {
                    "a:camera": { _attr: { prst: "perspectiveFront" } },
                },
                {
                    "a:lightRig": [
                        { _attr: { rig: "soft", dir: "r" } },
                        { "a:rot": { _attr: { lat: 0, lon: 0, rev: 2700000 } } },
                    ],
                },
            ],
        });
    });

    it("should create a scene3d with backdrop", () => {
        const tree = new Formatter().format(
            createScene3D({
                camera: { preset: "perspectiveFront" },
                lightRig: { rig: "threePt", direction: "t" },
                backdrop: {
                    anchor: { x: 0, y: 0, z: 0 },
                    normal: { dx: 0, dy: 0, dz: 1 },
                    up: { dx: 0, dy: 1, dz: 0 },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scene3d": [
                {
                    "a:camera": { _attr: { prst: "perspectiveFront" } },
                },
                {
                    "a:lightRig": { _attr: { rig: "threePt", dir: "t" } },
                },
                {
                    "a:backdrop": [
                        { "a:anchor": { _attr: { x: 0, y: 0, z: 0 } } },
                        { "a:norm": { _attr: { dx: 0, dy: 0, dz: 1 } } },
                        { "a:up": { _attr: { dx: 0, dy: 1, dz: 0 } } },
                    ],
                },
            ],
        });
    });

    it("should create a scene3d with all options", () => {
        const tree = new Formatter().format(
            createScene3D({
                camera: {
                    preset: "isometricTopUp",
                    fov: 5400000,
                    zoom: "100%",
                    rotation: { lat: 0, lon: 0, rev: 5400000 },
                },
                lightRig: {
                    rig: "balanced",
                    direction: "tl",
                    rotation: { lat: 0, lon: 0, rev: 2700000 },
                },
                backdrop: {
                    anchor: { x: 100, y: 200, z: 300 },
                    normal: { dx: 0, dy: 0, dz: 1 },
                    up: { dx: 0, dy: 1, dz: 0 },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scene3d": [
                {
                    "a:camera": [
                        { _attr: { prst: "isometricTopUp", fov: 5400000, zoom: "100%" } },
                        { "a:rot": { _attr: { lat: 0, lon: 0, rev: 5400000 } } },
                    ],
                },
                {
                    "a:lightRig": [
                        { _attr: { rig: "balanced", dir: "tl" } },
                        { "a:rot": { _attr: { lat: 0, lon: 0, rev: 2700000 } } },
                    ],
                },
                {
                    "a:backdrop": [
                        { "a:anchor": { _attr: { x: 100, y: 200, z: 300 } } },
                        { "a:norm": { _attr: { dx: 0, dy: 0, dz: 1 } } },
                        { "a:up": { _attr: { dx: 0, dy: 1, dz: 0 } } },
                    ],
                },
            ],
        });
    });
});
