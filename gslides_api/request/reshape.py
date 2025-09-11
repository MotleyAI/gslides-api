from gslides_api.domain.domain import PageElementProperties, Transform
from gslides_api.domain.request import ApplyMode
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.request.request import UpdatePageElementTransformRequest


def reshape_like_request(
    target_shape: PageElementProperties,
    current_shape: PageElementProperties,
    object_id: str | None = None,
) -> list[GSlidesAPIRequest]:
    old_size = target_shape.absolute_size(units="in")
    new_size = current_shape.absolute_size(units="in")

    print(target_shape, "\n", old_size)
    print(current_shape, "\n", new_size)

    new_scaleX = old_size[0] / new_size[0]
    new_scaleY = old_size[1] / new_size[1]

    transform2 = Transform(
        translateX=target_shape.transform.translateX,
        translateY=target_shape.transform.translateY,
        scaleX=current_shape.transform.scaleX * new_scaleX,
        scaleY=current_shape.transform.scaleY * new_scaleY,
        unit=target_shape.transform.unit,
    )
    requests = [
        UpdatePageElementTransformRequest(
            objectId=object_id,
            transform=transform2.to_affine_transform(),
            applyMode=ApplyMode.ABSOLUTE,
        ),
    ]
    return requests
