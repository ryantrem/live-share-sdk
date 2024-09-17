/* eslint-disable react/no-unknown-property */
import { FC, Suspense, useEffect, useRef } from "react";
import { ArcRotateCamera, Scene } from "@babylonjs/core";
import { registerBuiltInLoaders } from "@babylonjs/loaders/dynamic";
import ErrorBoundary from "./ErrorBoundary";
import type { HTML3DElement } from "@babylonjs/viewer";
import "@babylonjs/viewer";

registerBuiltInLoaders();

interface HTML3DElementAttributes
    extends React.DetailedHTMLProps<
        React.HTMLAttributes<HTMLElement>,
        HTMLElement
    > {
    src?: string;
    env?: string;
}

declare global {
    namespace JSX {
        interface IntrinsicElements {
            "babylon-viewer": HTML3DElementAttributes;
        }
    }
}

interface IModelSceneViewerProps {
    modelFileName: string;
    onReadyObservable: (scene: Scene, camera: ArcRotateCamera) => void;
}

/**
 * Babylon JS 3D viewer
 */
export const ModelViewerScene: FC<IModelSceneViewerProps> = ({
    modelFileName,
    onReadyObservable,
}) => {
    const viewerRef = useRef<HTML3DElement>(null);

    useEffect(() => {
        const viewerElement = viewerRef.current;
        const abortController = new AbortController();
        if (viewerElement) {
            let camera: ArcRotateCamera;
            viewerElement.addEventListener(
                "viewerready",
                (event: CustomEvent) => {
                    const scene: Scene = event.detail.scene;
                    camera = scene.activeCamera as ArcRotateCamera;
                    onReadyObservable(scene, camera);
                },
                {
                    once: true,
                    signal: abortController.signal,
                }
            );
            viewerElement.addEventListener("modelchange", () => {
                if (camera) {
                    camera.useAutoRotationBehavior = false;
                }
            });
        }

        return () => {
            abortController.abort();
        };
    }, []);

    return (
        <ErrorBoundary fallback={<></>}>
            <Suspense fallback={<></>}>
                <babylon-viewer
                    ref={viewerRef}
                    style={{ width: "100%", height: "100%" }}
                    src={modelFileName}
                    env="https://assets.babylonjs.com/environments/ulmerMuenster.env"
                ></babylon-viewer>
            </Suspense>
        </ErrorBoundary>
    );
};
