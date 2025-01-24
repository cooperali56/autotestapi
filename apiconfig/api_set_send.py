from abc import ABC, abstractmethod


class ApiSetSend(ABC):
    """
    接口发送
    """

    @abstractmethod
    def _get_send(self):
        """
        发送get请求

        Returns:
            None
        """
        pass

    @abstractmethod
    def _post_send(self):
        """
        发送post请求

        Returns:
            None
        """
        pass


    @abstractmethod
    def _delete_send(self):
        """
        删除请

        Returns:
            None
        """
        pass


