/* eslint-disable react/no-unknown-property */
import {
    FollowModeType,
    LiveDataObjectInitializeState,
    TestLiveShareHost,
    UserMeetingRole,
} from "@microsoft/live-share";
import {
    LiveShareProvider,
    useLiveFollowMode,
    useLiveShareContext,
    useSharedMap,
} from "@microsoft/live-share-react";
import { FrameContexts, LiveShareHost } from "@microsoft/teams-js";
import { inTeams } from "../utils/inTeams";
import { FC, useState, useEffect, useRef, useCallback, ReactNode } from "react";
import { Vector3, Color3 } from "@babylonjs/core/Maths/math";
import "@babylonjs/loaders/glTF";
import {
    Nullable,
    PickingInfo,
    Scene as BabyScene,
    ArcRotateCamera,
    Scene,
} from "@babylonjs/core";
import { PBRMaterial } from "@babylonjs/core/Materials";
import { IPointerEvent } from "@babylonjs/core/Events";
import { HexColorPicker } from "react-colorful";
import { Button, Spinner, Text, tokens } from "@fluentui/react-components";
import {
    DecorativeOutline,
    FlexColumn,
    LiveAvatars,
    FollowModeInfoText,
    FollowModeSmallButton,
    FollowModeInfoBar,
    TopHeaderBar,
    ModelViewerScene,
    SingleUserModelViewer,
} from "../components";
import { LiveCanvasOverlay } from "../components/LiveCanvasOverlay";
import { isLiveShareSupported } from "../utils/teams-utils";
import { useSharingStatus } from "../hooks/useSharingStatus";

const IN_TEAMS = inTeams();

export const TabContent: FC = () => {
    const [isSupported] = useState(IN_TEAMS ? isLiveShareSupported() : true);

    if (!isSupported) {
        return <SingleUserModelViewer />;
    }
    return <LiveShareContentWrapper />;
};

const LiveShareContentWrapper: FC = () => {
    const [host] = useState(
        IN_TEAMS ? LiveShareHost.create() : TestLiveShareHost.create()
    );
    return (
        <LiveShareProvider joinOnLoad host={host}>
            <LoadingErrorWrapper>
                <LiveObjectViewer />
            </LoadingErrorWrapper>
        </LiveShareProvider>
    );
};

const LoadingErrorWrapper: FC<{
    children?: ReactNode;
}> = ({ children }) => {
    const { joined, joinError } = useLiveShareContext();
    if (joinError) {
        return <Text>{joinError?.message}</Text>;
    }
    if (!joined) {
        return (
            <FlexColumn fill="view" vAlign="center" hAlign="center">
                <Spinner />
            </FlexColumn>
        );
    }
    return <>{children}</>;
};

export interface ICustomFollowData {
    cameraPosition?: {
        x: number;
        y: number;
        z: number;
    };
}

export const ALLOWED_ROLES = [
    UserMeetingRole.organizer,
    UserMeetingRole.presenter,
];

/**
 * Component that uses several Live Share features for viewing a Babylon.js 3D model collaboratively.
 * useSharedMap is used to synchronize the color of the model.
 * useLiveFollowMode is used to enable following/presenting.
 * useLiveCanvas is used to enable synchronized pen/highlighter/cursors atop the model when in follow mode.
 */
const LiveObjectViewer: FC = () => {
    // Babylon scene reference
    const sceneRef = useRef<Nullable<BabyScene>>(null);
    // Babylon arc rotation camera reference
    const [camera, setCamera] = useState<Nullable<ArcRotateCamera>>(null);
    // Pointer reference for mouse inputs, which is used so cursors continue working while interacting with the 3D model
    const pointerElementRef = useRef<HTMLDivElement>(null);
    // Boolean tracking whether an incoming camera update is being applied
    const isApplyingRemoteCameraUpdate = useRef(false);

    /**
     * Synchronized SharedMap for the color values that correspond to a material in the loaded .glb file
     */
    const {
        map: colorsMap,
        setEntry: setMaterialColor,
        sharedMap: sharedColorsMap,
    } = useSharedMap("COLORS");
    /**
     * Selected material for the color picker UI
     */
    const [selectedMaterialName, setSelectedMaterialName] = useState<
        string | null
    >(null);

    /**
     * Following state to track which camera position to display
     */
    const {
        allUsers, // List of users with info about who they are following and their custom state value.
        state: remoteCameraState, // The relevant state based on who is presenting / the user is following
        update: updateUserCameraState, // Update the local user's state value
        liveFollowMode,
        startPresenting, // Start presenting / take control
        stopPresenting, // Release control
        followUser, // Start following a specific user
        stopFollowing, // Stop following the currently followed user
        beginSuspension, // Temporarily suspend following the presenter / followed user
        endSuspension, // Resume following the presenter / followed user
    } = useLiveFollowMode<ICustomFollowData>(
        "FOLLOW_MODE", // unique key for DDS
        // Initial value, can either be the value itself or a callback to get the value.
        () => {
            // We use a callback because the camera position may change by the time LiveFollowMode is initialized
            return {};
        }, // default value
        ALLOWED_ROLES // roles who can "take control" of presenting
    );

    /**
     * Callback for when the local user selected a new color to apply to the 3D model
     */
    const onChangeColor = useCallback(
        (value: string) => {
            if (!selectedMaterialName) return;
            if (!sceneRef.current) return;
            try {
                setMaterialColor(selectedMaterialName, value);
                const material =
                    sceneRef.current.getMaterialByName(selectedMaterialName);
                if (material && material instanceof PBRMaterial) {
                    const color = Color3.FromHexString(value);
                    material.albedoColor = color;
                }
            } catch (err: any) {
                console.error(err);
            }
        },
        [selectedMaterialName]
    );

    /**
     * Callback for when the user clicks on the scene
     */
    const handlePointerDown = useCallback(
        (evt: IPointerEvent) => {
            if (!sceneRef.current) return;

            const pickResult: Nullable<PickingInfo> = sceneRef.current.pick(
                evt.clientX,
                evt.clientY
            );

            if (pickResult && pickResult.hit && pickResult.pickedMesh) {
                const mesh = pickResult.pickedMesh;
                if (!mesh.material) return;
                // When the user clicks on a specific material in our object, we set it as selected to show the color picker
                setSelectedMaterialName(mesh.material.name);
                return;
            }
            if (!selectedMaterialName) return;
            setSelectedMaterialName(null);
        },
        [selectedMaterialName]
    );

    /**
     * Setup the onPointerDown event listener
     */
    useEffect(() => {
        if (sceneRef.current) {
            sceneRef.current.onPointerDown = handlePointerDown;
        }
        return () => {
            if (sceneRef.current) {
                sceneRef.current.onPointerDown = undefined;
            }
        };
    }, [handlePointerDown]);

    /**
     * Callback to update the material colors for the latest remote values
     */
    const applyRemoteColors = useCallback(() => {
        colorsMap.forEach((value, key) => {
            if (!sceneRef.current) return;
            const material = sceneRef.current.getMaterialByName(key);
            if (material && material instanceof PBRMaterial) {
                const color = Color3.FromHexString(value);
                material.albedoColor = color;
            }
        });
    }, [colorsMap]);

    /**
     * When the synchronized colorsMap value changes we apply it to our scene.
     */
    useEffect(() => {
        applyRemoteColors();
    }, [applyRemoteColors]);

    /**
     * Send camera position for local user to remote
     */
    const sendCameraPos = useCallback(() => {
        if (!camera) return;
        const cameraPosition = camera.position;
        updateUserCameraState({
            cameraPosition: {
                x: cameraPosition.x,
                y: cameraPosition.y,
                z: cameraPosition.z,
            },
        });
    }, [camera, updateUserCameraState]);

    /**
     * Callback to snap camera position to presenting user when the remote value changes
     */
    const snapCameraIfFollowingUser = useCallback(() => {
        if (!remoteCameraState?.value.cameraPosition) return;
        // We do not need to snap to a remote value when referencing the local user's value
        if (remoteCameraState.isLocalValue) return;
        if (!camera) return;
        const remoteCameraPos = new Vector3(
            remoteCameraState.value.cameraPosition.x,
            remoteCameraState.value.cameraPosition.y,
            remoteCameraState.value.cameraPosition.z
        );
        if (sceneRef.current) {
            sceneRef.current.onBeforeRenderObservable.addOnce(() => {
                isApplyingRemoteCameraUpdate.current = true;
                camera.setPosition(remoteCameraPos);
                // Update the camera now (don't wait for the next render)
                camera.update();
                camera.updateCache();
                isApplyingRemoteCameraUpdate.current = false;
            });
        }
    }, [remoteCameraState]);

    /**
     * Update camera position when following user's presence changes
     */
    useEffect(() => {
        snapCameraIfFollowingUser();
    }, [snapCameraIfFollowingUser]);

    /**
     * Callback when the arcLightCamera position changes
     */
    const onCameraViewMatrixChanged = useCallback(() => {
        // If we are not applying a remote camera update, then it must be a local update
        if (!isApplyingRemoteCameraUpdate.current) {
            sendCameraPos();
            // The user selected a camera position that is not in sync with the remote position, so we start a new suspension.
            // The user will be able to return in sync with the remote position when `endSuspension` is called.
            beginSuspension();
        }
    }, [sendCameraPos, beginSuspension]);

    /**
     * Set/update the camera view matrix change listener
     */
    useEffect(() => {
        if (!camera) return;
        // Add an observable
        const observer = camera.onViewMatrixChangedObservable.add(
            onCameraViewMatrixChanged
        );
        // Clear observables on unmount
        return () => {
            observer.remove();
        };
    }, [camera, onCameraViewMatrixChanged]);

    // Choose to only get sharing status in meetingStage.
    // We use this to take control on meeting stage if isShareInitiator == true on first load.
    const sharingStatus = useSharingStatus(FrameContexts.meetingStage);
    // Start presenting if nobody is in control and local user isShareInitiator (meetings only)
    const hasInitiallyPresentedRef = useRef<boolean>(false);
    useEffect(() => {
        if (
            liveFollowMode?.initializeState !==
            LiveDataObjectInitializeState.succeeded
        )
            return;
        if (!sharingStatus?.isShareInitiator) return;
        if (!remoteCameraState) return;
        if (remoteCameraState.type !== FollowModeType.local) return;
        if (hasInitiallyPresentedRef.current) return;
        hasInitiallyPresentedRef.current = true;
        startPresenting();
    }, [
        sharingStatus?.isShareInitiator,
        remoteCameraState?.type,
        liveFollowMode?.initializeState,
        startPresenting,
    ]);

    // Show LiveCanvas overlay and/or decorative overlay while in follow mode
    const followModeActive = remoteCameraState
        ? [
              FollowModeType.followPresenter,
              FollowModeType.followUser,
              FollowModeType.activePresenter,
              FollowModeType.activeFollowers,
          ].includes(remoteCameraState.type)
        : false;

    return (
        <>
            {!!remoteCameraState && (
                <TopHeaderBar
                    right={
                        <>
                            {(remoteCameraState.type === FollowModeType.local ||
                                remoteCameraState.type ===
                                    FollowModeType.activeFollowers) && (
                                <Button
                                    appearance="primary"
                                    onClick={startPresenting}
                                >
                                    {"Spotlight me"}
                                </Button>
                            )}
                            {remoteCameraState.type ===
                                FollowModeType.activePresenter && (
                                <Button
                                    appearance="primary"
                                    onClick={stopPresenting}
                                >
                                    {"Stop spotlight"}
                                </Button>
                            )}
                            {(remoteCameraState.type ===
                                FollowModeType.suspendFollowUser ||
                                remoteCameraState.type ===
                                    FollowModeType.followUser) && (
                                <Button
                                    appearance="primary"
                                    onClick={stopFollowing}
                                >
                                    {"Stop following"}
                                </Button>
                            )}
                            {(remoteCameraState.type ===
                                FollowModeType.followPresenter ||
                                remoteCameraState.type ===
                                    FollowModeType.suspendFollowPresenter) && (
                                <Button
                                    appearance="secondary"
                                    onClick={startPresenting}
                                >
                                    {"Take control"}
                                </Button>
                            )}
                        </>
                    }
                >
                    <LiveAvatars
                        allUsers={allUsers}
                        remoteCameraState={remoteCameraState}
                        onFollowUser={followUser}
                    />
                </TopHeaderBar>
            )}
            <FlexColumn fill="view" ref={pointerElementRef}>
                <ModelViewerScene
                    modelFileName="https://raw.githubusercontent.com/BabylonJS/Assets/master/splats/gs_Skull.splat"
                    onReadyObservable={(
                        scene: Scene,
                        camera: ArcRotateCamera
                    ) => {
                        sceneRef.current = scene;
                        sceneRef.current.onPointerDown = handlePointerDown;
                        applyRemoteColors();
                        snapCameraIfFollowingUser();
                        setCamera(camera);
                    }}
                />
            </FlexColumn>
            {/* Decorative border while following / presenting */}
            {!!remoteCameraState && followModeActive && (
                <DecorativeOutline
                    borderColor={
                        remoteCameraState.type ===
                            FollowModeType.activePresenter ||
                        remoteCameraState.type ===
                            FollowModeType.activeFollowers
                            ? tokens.colorPaletteRedBackground3
                            : tokens.colorPaletteBlueBorderActive
                    }
                />
            )}
            {!!sharedColorsMap && !!selectedMaterialName && (
                <HexColorPicker
                    color={
                        selectedMaterialName
                            ? colorsMap.get(selectedMaterialName)
                            : undefined
                    }
                    onChange={onChangeColor}
                    style={{
                        position: "absolute",
                        top: "72px",
                        right: "24px",
                    }}
                />
            )}
            {/* LiveCanvas for inking */}
            {!!remoteCameraState && followModeActive && (
                <LiveCanvasOverlay
                    pointerElementRef={pointerElementRef}
                    followingUserId={remoteCameraState.followingUserId}
                    zPosition={remoteCameraState.value?.cameraPosition?.z ?? 0}
                />
            )}
            {/* Follow mode information / actions */}
            {!!remoteCameraState &&
                remoteCameraState.type !== FollowModeType.local && (
                    <FollowModeInfoBar remoteCameraState={remoteCameraState}>
                        <FollowModeInfoText />
                        {remoteCameraState.type ===
                            FollowModeType.activePresenter && (
                            <FollowModeSmallButton onClick={stopPresenting}>
                                {"STOP"}
                            </FollowModeSmallButton>
                        )}
                        {remoteCameraState.type ===
                            FollowModeType.followUser && (
                            <FollowModeSmallButton onClick={stopFollowing}>
                                {"STOP"}
                            </FollowModeSmallButton>
                        )}
                        {remoteCameraState.type ===
                            FollowModeType.suspendFollowPresenter && (
                            <FollowModeSmallButton onClick={endSuspension}>
                                {"FOLLOW"}
                            </FollowModeSmallButton>
                        )}
                        {remoteCameraState.type ===
                            FollowModeType.suspendFollowUser && (
                            <FollowModeSmallButton onClick={endSuspension}>
                                {"RESUME"}
                            </FollowModeSmallButton>
                        )}
                    </FollowModeInfoBar>
                )}
        </>
    );
};
